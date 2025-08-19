import requests
import json
from config import Config
import logging

logger = logging.getLogger(__name__)

class OpenAPIClient:
    """Client for making requests to OpenAPI services"""
    
    def __init__(self):
        self.provider = Config.OPENAI_PROVIDER
        
        if self.provider == 'azure':
            self.api_key = Config.AZURE_OPENAI_API_KEY
            self.api_base = Config.AZURE_OPENAI_ENDPOINT
            self.deployment = Config.AZURE_OPENAI_DEPLOYMENT
            self.api_version = Config.AZURE_OPENAI_API_VERSION
            
            if not all([self.api_key, self.api_base, self.deployment]):
                raise ValueError("Azure OpenAI configuration incomplete")
        else:
            self.api_key = Config.OPENAI_API_KEY
            self.api_base = Config.OPENAI_API_BASE
            self.model = Config.OPENAI_MODEL
            
            if not self.api_key:
                raise ValueError("OPENAI_API_KEY not configured")
    
    def make_chat_completion(self, messages, temperature=0.7, max_tokens=1000):
        """
        Make a chat completion request to OpenAI API
        
        Args:
            messages (list): List of message dictionaries
            temperature (float): Sampling temperature
            max_tokens (int): Maximum tokens to generate
            
        Returns:
            dict: API response
        """
        if self.provider == 'azure':
            url = f"{self.api_base}/openai/deployments/{self.deployment}/chat/completions"
            params = {'api-version': self.api_version}
        else:
            url = f"{self.api_base}/chat/completions"
            params = {}
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens
        }
        
        # Add model for standard OpenAI (not needed for Azure)
        if self.provider != 'azure':
            data["model"] = self.model
        
        try:
            logger.info(f"Making OpenAI API request to {url}")
            response = requests.post(url, headers=headers, json=data, params=params, timeout=120)
            response.raise_for_status()
            
            result = response.json()
            logger.info("OpenAI API request successful")
            return result
            
        except requests.exceptions.RequestException as e:
            logger.error(f"OpenAI API request failed: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"Response status: {e.response.status_code}")
                logger.error(f"Response body: {e.response.text}")
            raise
    
    def make_completion(self, prompt, temperature=0.7, max_tokens=1000):
        """
        Make a completion request to OpenAI API (legacy endpoint)
        
        Args:
            prompt (str): The prompt text
            temperature (float): Sampling temperature
            max_tokens (int): Maximum tokens to generate
            
        Returns:
            dict: API response
        """
        if self.provider == 'azure':
            url = f"{self.api_base}/openai/deployments/{self.deployment}/completions"
            params = {'api-version': self.api_version}
        else:
            url = f"{self.api_base}/completions"
            params = {}
        
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "prompt": prompt,
            "temperature": temperature,
            "max_tokens": max_tokens
        }
        
        # Add model for standard OpenAI (not needed for Azure)
        if self.provider != 'azure':
            data["model"] = "text-davinci-003"  # Legacy model
        
        try:
            logger.info(f"Making OpenAI completion request to {url}")
            response = requests.post(url, headers=headers, json=data, params=params, timeout=120)
            response.raise_for_status()
            
            result = response.json()
            logger.info("OpenAI completion request successful")
            return result
            
        except requests.exceptions.RequestException as e:
            logger.error(f"OpenAI completion request failed: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                logger.error(f"Response status: {e.response.status_code}")
                logger.error(f"Response body: {e.response.text}")
            raise

class AnthropicClient:
    """Client for making requests to Anthropic API"""
    
    def __init__(self):
        self.api_key = Config.ANTHROPIC_API_KEY
        
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY not configured")
    
    def make_message(self, messages, model="claude-3-sonnet-20240229", max_tokens=1000):
        """
        Make a message request to Anthropic API
        
        Args:
            messages (list): List of message dictionaries
            model (str): Model to use
            max_tokens (int): Maximum tokens to generate
            
        Returns:
            dict: API response
        """
        url = "https://api.anthropic.com/v1/messages"
        
        headers = {
            "x-api-key": self.api_key,
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01"
        }
        
        data = {
            "model": model,
            "messages": messages,
            "max_tokens": max_tokens
        }
        
        try:
            logger.info(f"Making Anthropic API request to {url}")
            response = requests.post(url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            
            result = response.json()
            logger.info("Anthropic API request successful")
            return result
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Anthropic API request failed: {str(e)}")
            raise 