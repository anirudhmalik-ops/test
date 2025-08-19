# Postman Collections for File AI Formatter API

This directory contains Postman collections and environment files for testing the File AI Formatter API.

## Files Included

1. **`File_AI_Formatter_API.postman_collection.json`** - Basic collection with all endpoints
2. **`File_AI_Formatter_API_Enhanced.postman_collection.json`** - Enhanced collection with automated tests and workflow
3. **`File_AI_Formatter_Environment.postman_environment.json`** - Environment variables for different deployment scenarios

## Setup Instructions

### 1. Import Collections

1. Open Postman
2. Click **Import** button
3. Import the following files:
   - `File_AI_Formatter_API.postman_collection.json`
   - `File_AI_Formatter_API_Enhanced.postman_collection.json`

### 2. Import Environment

1. Click **Import** button
2. Import `File_AI_Formatter_Environment.postman_environment.json`
3. Select the environment from the dropdown in the top-right corner

### 3. Configure Environment Variables

The environment includes these variables:

| Variable | Default Value | Description |
|----------|---------------|-------------|
| `base_url` | `http://localhost:5001` | Base URL for the API server |
| `processed_filename` | (empty) | Automatically set from upload response |
| `download_url` | (empty) | Automatically set from upload response |

## Collections Overview

### Basic Collection (`File_AI_Formatter_API.postman_collection.json`)

Simple collection with all API endpoints:

- **Health Check** - Verify API is running
- **API Status** - Check configuration status
- **Upload Excel File** - Upload and process Excel files
- **Download Processed File** - Download processed files
- **OpenAI Chat** - Test OpenAI integration
- **Anthropic Chat** - Test Anthropic integration

### Enhanced Collection (`File_AI_Formatter_API_Enhanced.postman_collection.json`)

Advanced collection with automated tests and workflow:

- **Automated Tests** - Each request includes validation tests
- **Workflow Automation** - Automatic file handling between requests
- **Response Validation** - Comprehensive response checking
- **Error Handling** - Graceful handling of expected errors

## Usage Workflow

### Complete Excel Processing Workflow

1. **Health Check** - Verify API is running
2. **API Status** - Confirm services are configured
3. **Upload Excel File** - Upload your Excel file for processing
4. **Download Processed File** - Download the AI-processed result

### Testing Individual Endpoints

You can test each endpoint independently:

- **Health Check**: `GET {{base_url}}/health`
- **API Status**: `GET {{base_url}}/api/status`
- **OpenAI Chat**: `POST {{base_url}}/api/openai/chat`
- **Anthropic Chat**: `POST {{base_url}}/api/anthropic/chat`

## File Upload Instructions

### For Excel File Upload:

1. Select the **"Upload Excel File"** request
2. In the **Body** tab, select **form-data**
3. Click **Select Files** next to the `file` field
4. Choose your Excel file (.xlsx or .xls format)
5. Click **Send**

### Expected Response:

```json
{
  "message": "File processed successfully",
  "output_file": "processed_20240115_143022.xlsx",
  "download_url": "/api/download/processed_20240115_143022.xlsx"
}
```

### Download Processed File:

1. The enhanced collection automatically stores the filename
2. Use the **"Download Processed File"** request
3. The filename is automatically populated from the upload response

## Automated Tests

The enhanced collection includes automated tests for each endpoint:

### Health Check Tests:
- Status code is 200
- Response has correct message
- Response time is under 1000ms

### API Status Tests:
- Status code is 200
- Required fields are present
- OpenAI is configured
- Excel processor is configured

### Upload Tests:
- Status code is 200
- Response has required fields
- Success message is correct
- Automatically stores filename for download

### Download Tests:
- Status code is 200
- Response is an Excel file
- File has content

### Chat Tests:
- Status code is 200
- Response has correct structure
- Content is present

## Environment Configurations

### Local Development
```json
{
  "base_url": "http://localhost:5001"
}
```

### Staging Environment
```json
{
  "base_url": "https://staging-api.yourdomain.com"
}
```

### Production Environment
```json
{
  "base_url": "https://api.yourdomain.com"
}
```

## Troubleshooting

### Common Issues

1. **Connection Refused**
   - Ensure the Flask server is running
   - Check the `base_url` variable
   - Verify the port (default: 5001)

2. **File Upload Fails**
   - Check file size (max 16MB)
   - Ensure file is .xlsx or .xls format
   - Verify file is not corrupted

3. **AI Processing Fails**
   - Check API keys in `.env` file
   - Verify Azure OpenAI configuration
   - Check logs for detailed error messages

4. **Download Fails**
   - Ensure upload was successful
   - Check if filename is correctly set
   - Verify file exists on server

### Debug Mode

Enable debug mode in Postman:
1. Go to **Settings** (gear icon)
2. Enable **Show response headers**
3. Enable **Show response size**
4. Enable **Show response time**

## API Documentation

### Endpoints Summary

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/health` | Health check |
| GET | `/api/status` | API configuration status |
| POST | `/api/upload/excel` | Upload Excel file for processing |
| GET | `/api/download/{filename}` | Download processed file |
| POST | `/api/openai/chat` | OpenAI chat completion |
| POST | `/api/anthropic/chat` | Anthropic chat completion |

### Request/Response Examples

#### Health Check
```http
GET http://localhost:5001/health
```

Response:
```json
{
  "message": "project running"
}
```

#### API Status
```http
GET http://localhost:5001/api/status
```

Response:
```json
{
  "openai_configured": true,
  "openai_provider": "azure",
  "anthropic_configured": false,
  "excel_processor_configured": true,
  "missing_keys": [],
  "azure_config": {
    "endpoint": "https://pavingptai.openai.azure.com",
    "deployment": "gpt-4o",
    "api_version": "2024-02-15-preview"
  }
}
```

#### OpenAI Chat
```http
POST http://localhost:5001/api/openai/chat
Content-Type: application/json

{
  "messages": [
    {
      "role": "user",
      "content": "Hello, how are you?"
    }
  ],
  "temperature": 0.7,
  "max_tokens": 1000
}
```

## Running the Complete Workflow

1. **Start the Flask server:**
   ```bash
   python app.py
   ```

2. **Import the collections into Postman**

3. **Run the enhanced collection in order:**
   - Health Check
   - API Status
   - Upload Excel File (with your file)
   - Download Processed File
   - OpenAI Chat Test
   - Anthropic Chat Test

4. **Check the test results** in the Postman console

## Support

For issues with the API:
- Check the Flask server logs
- Review the `.env` configuration
- Verify API keys are correct
- Check file permissions

For issues with Postman:
- Update to the latest version
- Clear cache and cookies
- Re-import collections
- Check environment variables 