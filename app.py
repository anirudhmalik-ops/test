from flask import Flask, jsonify, request, send_file, render_template
import logging
from logging.handlers import TimedRotatingFileHandler
import os
from datetime import datetime
from config import Config
from api_client import OpenAPIClient, AnthropicClient
from excel_processor import ExcelProcessor
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Create logs directory if it doesn't exist
if not os.path.exists('logs'):
    os.makedirs('logs')

# Create uploads directory if it doesn't exist
if not os.path.exists('uploads'):
    os.makedirs('uploads')

# Create templates directory if it doesn't exist
if not os.path.exists('templates'):
    os.makedirs('templates')

# Configure logging
def setup_logging():
    # Create a logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Create TimedRotatingFileHandler for daily log rotation
    log_file = os.path.join('logs', 'app.log')
    file_handler = TimedRotatingFileHandler(
        log_file,
        when='midnight',
        interval=1,
        backupCount=30,  # Keep 30 days of logs
        encoding='utf-8'
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    
    # Create console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# Setup logging
logger = setup_logging()

# Initialize API clients
try:
    openai_client = OpenAPIClient()
    logger.info(f"OpenAI client initialized successfully (provider: {Config.OPENAI_PROVIDER})")
except ValueError as e:
    logger.warning(f"OpenAI client not initialized: {e}")
    openai_client = None

try:
    anthropic_client = AnthropicClient()
    logger.info("Anthropic client initialized successfully")
except ValueError as e:
    logger.warning(f"Anthropic client not initialized: {e}")
    anthropic_client = None

# Initialize Excel processor
try:
    excel_processor = ExcelProcessor()
    logger.info("Excel processor initialized successfully")
except Exception as e:
    logger.warning(f"Excel processor not initialized: {e}")
    excel_processor = None

# File upload configuration
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """
    Serve the upload form
    """
    return render_template('upload.html')

@app.route('/health', methods=['GET'])
def health_check():
    """
    Health check endpoint that returns a 200 status with a JSON response
    """
    logger.info("Health check endpoint accessed")
    return jsonify({"message": "project running"}), 200

@app.route('/api/upload/excel', methods=['POST'])
def upload_excel():
    """
    Endpoint to upload and process Excel files
    """
    if not excel_processor:
        return jsonify({"error": "Excel processor not available"}), 500
    
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        
        # Check if file was selected
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        # Check file extension
        if not allowed_file(file.filename):
            return jsonify({"error": "Invalid file type. Only .xlsx and .xls files are allowed"}), 400
        
        # Check file size
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            return jsonify({"error": "File too large. Maximum size is 16MB"}), 400
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        uploaded_filename = f"{timestamp}_{filename}"
        uploaded_file_path = os.path.join("uploads", uploaded_filename)
        
        file.save(uploaded_file_path)
        logger.info(f"File uploaded: {uploaded_filename}")
        
        # Process the file
        logger.info(f"Starting file processing: {uploaded_filename}")
        output_path = excel_processor.process_uploaded_file(uploaded_file_path)
        
        # Get the filename for download
        output_filename = os.path.basename(output_path)
        
        # Clean up uploaded file
        try:
            os.remove(uploaded_file_path)
            logger.info(f"Cleaned up uploaded file: {uploaded_filename}")
        except Exception as e:
            logger.warning(f"Failed to clean up uploaded file: {e}")
        
        return jsonify({
            "message": "File processed successfully",
            "output_file": output_filename,
            "download_url": f"/api/download/{output_filename}"
        }), 200
        
    except Exception as e:
        logger.error(f"Error processing uploaded file: {str(e)}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    """
    Endpoint to download processed files
    """
    try:
        file_path = os.path.join("document", filename)
        
        if not os.path.exists(file_path):
            return jsonify({"error": "File not found"}), 404
        
        logger.info(f"File downloaded: {filename}")
        return send_file(file_path, as_attachment=True, download_name=filename)
        
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        return jsonify({"error": "Download failed"}), 500

@app.route('/api/openai/chat', methods=['POST'])
def openai_chat():
    """
    Endpoint to make OpenAI chat completion requests
    """
    if not openai_client:
        return jsonify({"error": "OpenAI API key not configured"}), 500
    
    try:
        data = request.get_json()
        
        if not data or 'messages' not in data:
            return jsonify({"error": "Messages are required"}), 400
        
        messages = data['messages']
        temperature = data.get('temperature', 0.7)
        max_tokens = data.get('max_tokens', 1000)
        
        logger.info(f"Making OpenAI chat request with {len(messages)} messages (provider: {Config.OPENAI_PROVIDER})")
        
        response = openai_client.make_chat_completion(
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens
        )
        
        return jsonify(response), 200
        
    except Exception as e:
        logger.error(f"OpenAI chat request failed: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/anthropic/chat', methods=['POST'])
def anthropic_chat():
    """
    Endpoint to make Anthropic chat requests
    """
    if not anthropic_client:
        return jsonify({"error": "Anthropic API key not configured"}), 500
    
    try:
        data = request.get_json()
        
        if not data or 'messages' not in data:
            return jsonify({"error": "Messages are required"}), 400
        
        messages = data['messages']
        model = data.get('model', 'claude-3-sonnet-20240229')
        max_tokens = data.get('max_tokens', 1000)
        
        logger.info(f"Making Anthropic chat request with {len(messages)} messages")
        
        response = anthropic_client.make_message(
            messages=messages,
            model=model,
            max_tokens=max_tokens
        )
        
        return jsonify(response), 200
        
    except Exception as e:
        logger.error(f"Anthropic chat request failed: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/api/status', methods=['GET'])
def api_status():
    """
    Endpoint to check API key configuration status
    """
    status = {
        "openai_configured": openai_client is not None,
        "openai_provider": Config.OPENAI_PROVIDER,
        "anthropic_configured": anthropic_client is not None,
        "excel_processor_configured": excel_processor is not None,
        "missing_keys": Config.validate_api_keys()
    }
    
    # Add provider-specific information
    if Config.OPENAI_PROVIDER == 'azure':
        status["azure_config"] = {
            "endpoint": Config.AZURE_OPENAI_ENDPOINT,
            "deployment": Config.AZURE_OPENAI_DEPLOYMENT,
            "api_version": Config.AZURE_OPENAI_API_VERSION
        }
    else:
        status["openai_config"] = {
            "api_base": Config.OPENAI_API_BASE,
            "model": Config.OPENAI_MODEL
        }
    
    return jsonify(status), 200

@app.before_request
def log_request():
    """Log all incoming requests"""
    logger.info(f"Request: {request.method} {request.url} from {request.remote_addr}")

@app.after_request
def log_response(response):
    """Log all outgoing responses"""
    logger.info(f"Response: {response.status_code} for {request.method} {request.url}")
    return response

if __name__ == '__main__':
    logger.info("Starting Flask application")
    app.run(debug=True, host='0.0.0.0', port=5001) 