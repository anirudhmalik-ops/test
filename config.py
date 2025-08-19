import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

class Config:
    """Configuration class for the application"""
    
    # Flask configuration
    SECRET_KEY = os.getenv('SECRET_KEY', 'your-secret-key-change-this')
    
    # OpenAI configuration (standard OpenAI)
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    OPENAI_API_BASE = os.getenv('OPENAI_API_BASE', 'https://api.openai.com/v1')
    OPENAI_MODEL = os.getenv('OPENAI_MODEL', 'gpt-3.5-turbo')
    
    # Azure OpenAI configuration
    AZURE_OPENAI_API_KEY = os.getenv('AZURE_OPENAI_API_KEY')
    AZURE_OPENAI_ENDPOINT = os.getenv('AZURE_OPENAI_ENDPOINT')
    AZURE_OPENAI_DEPLOYMENT = os.getenv('AZURE_OPENAI_DEPLOYMENT')
    AZURE_OPENAI_API_VERSION = os.getenv('AZURE_OPENAI_API_VERSION', '2024-02-15-preview')
    
    # API provider selection
    OPENAI_PROVIDER = os.getenv('OPENAI_PROVIDER', 'openai')  # 'openai' or 'azure'
    
    # Other API keys (add as needed)
    ANTHROPIC_API_KEY = os.getenv('ANTHROPIC_API_KEY')
    GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')
    
    @classmethod
    def validate_api_keys(cls):
        """Validate that required API keys are present"""
        missing_keys = []
        
        if cls.OPENAI_PROVIDER == 'azure':
            if not cls.AZURE_OPENAI_API_KEY:
                missing_keys.append('AZURE_OPENAI_API_KEY')
            if not cls.AZURE_OPENAI_ENDPOINT:
                missing_keys.append('AZURE_OPENAI_ENDPOINT')
            if not cls.AZURE_OPENAI_DEPLOYMENT:
                missing_keys.append('AZURE_OPENAI_DEPLOYMENT')
        else:
            if not cls.OPENAI_API_KEY:
                missing_keys.append('OPENAI_API_KEY')
        
        return missing_keys 