# File AI Formatter - Excel Processing API

A Python Flask application that processes Excel files using AI, applies formatting based on templates, and generates formatted output files. Supports both standard OpenAI and Azure OpenAI.

## Features

- ✅ **Excel File Upload**: Upload .xlsx and .xls files
- ✅ **AI Processing**: Uses OpenAI to analyze and process data
- ✅ **Template-Based Formatting**: Applies formatting based on predefined templates
- ✅ **File Download**: Download processed files
- ✅ **Comprehensive Logging**: Daily log rotation with detailed logging
- ✅ **Multiple AI Providers**: Support for OpenAI and Azure OpenAI
- ✅ **Web Interface**: Simple HTML upload form for testing

## Setup

1. **Activate the virtual environment:**
   ```bash
   source venv/bin/activate
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure API keys:**
   ```bash
   cp env.example .env
   ```
   
   Edit the `.env` file with your API keys:
   ```bash
   # For Azure OpenAI (recommended)
   OPENAI_PROVIDER=azure
   AZURE_OPENAI_API_KEY=your-azure-openai-api-key-here
   AZURE_OPENAI_ENDPOINT=https://your-resource-name.openai.azure.com
   AZURE_OPENAI_DEPLOYMENT=your-deployment-name
   
   # For standard OpenAI (alternative)
   OPENAI_PROVIDER=openai
   OPENAI_API_KEY=your-openai-api-key-here
   ```

## Running the Application

Start the Flask server:
```bash
python app.py
```

The server will start on `http://localhost:5001`

## Web Interface

Visit `http://localhost:5001` to access the web upload interface.

## API Endpoints

### Health Check
- **URL**: `/health`
- **Method**: `GET`
- **Response**: 
  - Status: 200
  - Body: `{"message": "project running"}`

### API Status
- **URL**: `/api/status`
- **Method**: `GET`
- **Response**: Shows which API keys are configured and provider information

### Excel File Upload
- **URL**: `/api/upload/excel`
- **Method**: `POST`
- **Content-Type**: `multipart/form-data`
- **Body**: File upload with field name `file`
- **Response**: 
  ```json
  {
    "message": "File processed successfully",
    "output_file": "processed_20240115_143022.xlsx",
    "download_url": "/api/download/processed_20240115_143022.xlsx"
  }
  ```

### File Download
- **URL**: `/api/download/<filename>`
- **Method**: `GET`
- **Response**: File download

### OpenAI Chat
- **URL**: `/api/openai/chat`
- **Method**: `POST`
- **Body**:
  ```json
  {
    "messages": [
      {"role": "user", "content": "Hello, how are you?"}
    ],
    "temperature": 0.7,
    "max_tokens": 1000
  }
  ```

### Anthropic Chat
- **URL**: `/api/anthropic/chat`
- **Method**: `POST`
- **Body**:
  ```json
  {
    "messages": [
      {"role": "user", "content": "Hello, how are you?"}
    ],
    "model": "claude-3-sonnet-20240229",
    "max_tokens": 1000
  }
  ```

## Excel Processing Workflow

1. **Upload**: User uploads an Excel file (.xlsx or .xls)
2. **Analysis**: System reads the uploaded file and analyzes the template structure
3. **AI Processing**: OpenAI processes the data and maps it to the template format
4. **Formatting**: System applies the template formatting to the processed data
5. **Output**: Generated file is saved and made available for download

## File Structure

```
file-ai-formatter/
├── app.py                 # Main Flask application
├── config.py             # Configuration management
├── api_client.py         # OpenAI and Anthropic API clients
├── excel_processor.py    # Excel file processing logic
├── requirements.txt      # Python dependencies
├── .env                  # Environment variables (create from env.example)
├── .gitignore           # Git ignore rules
├── document/            # Template and output files
│   ├── Template-Waaree.xlsx          # Template file
│   └── processed_*.xlsx              # Generated output files
├── uploads/             # Temporary uploaded files
├── logs/                # Application logs
│   └── app.log          # Current log file
└── templates/           # HTML templates
    └── upload.html      # Upload form interface
```

## API Key Configuration

The application supports multiple API providers:

### Azure OpenAI (Recommended)
- **Required**: `AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_ENDPOINT`, `AZURE_OPENAI_DEPLOYMENT`
- **Optional**: `AZURE_OPENAI_API_VERSION`
- **Provider Setting**: `OPENAI_PROVIDER=azure`

### Standard OpenAI
- **Required**: `OPENAI_API_KEY`
- **Optional**: `OPENAI_API_BASE`, `OPENAI_MODEL`
- **Provider Setting**: `OPENAI_PROVIDER=openai`

### Anthropic
- **Required**: `ANTHROPIC_API_KEY`

## Testing the API

### Using the Web Interface
1. Visit `http://localhost:5001`
2. Select an Excel file
3. Click "Process File"
4. Download the processed file

### Using curl

#### Upload Excel File
```bash
curl -X POST http://localhost:5001/api/upload/excel \
  -F "file=@your_file.xlsx"
```

#### Check API Status
```bash
curl http://localhost:5001/api/status
```

#### Health Check
```bash
curl http://localhost:5001/health
```

#### OpenAI Chat
```bash
curl -X POST http://localhost:5001/api/openai/chat \
  -H "Content-Type: application/json" \
  -d '{
    "messages": [
      {"role": "user", "content": "Hello, how are you?"}
    ]
  }'
```

## Logging

The application includes comprehensive logging:

- **Log Directory**: `logs/`
- **Current Log File**: `logs/app.log`
- **Daily Rotation**: Log files are automatically rotated at midnight
- **Log Retention**: Keeps 30 days of log files
- **Log Format**: `timestamp - logger_name - level - message`

### What Gets Logged
- Application startup
- All incoming HTTP requests
- All outgoing HTTP responses
- File uploads and processing
- AI API requests and responses
- Error messages and exceptions

## Viewing Logs

To view the current day's logs:
```bash
tail -f logs/app.log
```

To view logs from a specific date:
```bash
cat logs/app.log.2024-01-15
```

## File Upload Limits

- **Maximum File Size**: 16MB
- **Allowed Extensions**: .xlsx, .xls
- **Temporary Storage**: Files are stored in `uploads/` directory during processing
- **Cleanup**: Uploaded files are automatically cleaned up after processing

## Template Configuration

The system uses `document/Template-Waaree.xlsx` as the template file. This file defines:
- Sheet structure
- Column headers
- Formatting (fonts, colors, alignment)
- Data types

## Error Handling

The application includes comprehensive error handling:
- File validation (type, size)
- AI API error handling
- File processing error handling
- Detailed error logging
- User-friendly error messages

## Security Features

- File type validation
- File size limits
- Secure filename handling
- Environment variable configuration
- Input sanitization
- Comprehensive logging for monitoring

## Troubleshooting

### Common Issues

1. **File Upload Fails**
   - Check file size (max 16MB)
   - Ensure file is .xlsx or .xls format
   - Check logs for detailed error messages

2. **AI Processing Fails**
   - Verify API keys are configured correctly
   - Check API provider settings
   - Review logs for API error details

3. **Template Not Found**
   - Ensure `document/Template-Waaree.xlsx` exists
   - Check file permissions

4. **Download Fails**
   - Verify processed file exists in `document/` directory
   - Check file permissions

### Debug Mode

Run the application in debug mode for detailed error messages:
```bash
python app.py
```

## Development

### Adding New Features

1. **New API Endpoints**: Add routes in `app.py`
2. **New AI Providers**: Extend `api_client.py`
3. **New File Formats**: Extend `excel_processor.py`
4. **New Templates**: Add template files to `document/` directory

### Testing

1. **Unit Tests**: Add tests for individual components
2. **Integration Tests**: Test the complete workflow
3. **API Tests**: Test all endpoints with various inputs

## Deployment

### Production Considerations

1. **Environment Variables**: Set production API keys
2. **Logging**: Configure production logging levels
3. **File Storage**: Use cloud storage for uploaded files
4. **Security**: Enable HTTPS, add authentication
5. **Monitoring**: Add health checks and monitoring

### Docker Deployment

Create a `Dockerfile`:
```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 5001
CMD ["python", "app.py"]
```

## License

This project is licensed under the MIT License. 