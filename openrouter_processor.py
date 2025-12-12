#!/usr/bin/env python3
"""
OpenRouter API Processor for Note Taking App
Handles communication with OpenRouter API for AI-powered note processing.
"""

import os
import json
import time
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
import requests


@dataclass
class NoteProcessingConfig:
    """Configuration for note processing with OpenRouter API."""
    model: str = "openai/gpt-4o-mini"
    temperature: float = 0.3
    max_tokens: int = 4000
    pre_prompt: str = ""
    timeout: int = 120
    max_retries: int = 3
    retry_delay: float = 1.0


class OpenRouterProcessor:
    """Handles note processing using OpenRouter API."""
    
    def __init__(self, api_key: str = None, config: NoteProcessingConfig = None):
        self.config = config or NoteProcessingConfig()
        
        # Get API key from parameter, environment, or file
        self.api_key = api_key
        if not self.api_key:
            self.api_key = os.getenv("OPENROUTER_API_KEY")
        if not self.api_key:
            api_key_file = Path("openrouter_api_key.txt")
            if api_key_file.exists():
                try:
                    with open(api_key_file, 'r', encoding='utf-8') as f:
                        self.api_key = f.read().strip()
                except Exception as e:
                    logging.warning(f"Could not read API key from file: {e}")
        
        if not self.api_key:
            raise ValueError("OpenRouter API key not found. Please provide it via parameter, environment variable OPENROUTER_API_KEY, or openrouter_api_key.txt file.")
        
        # API configuration
        self.base_url = "https://openrouter.ai/api/v1"
        self.chat_url = f"{self.base_url}/chat/completions"
        self.models_url = f"{self.base_url}/models"
        
        # Headers for all requests
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "HTTP-Referer": "https://github.com/your-username/school-note-app",  # Optional: your app's URL
            "X-Title": "School Note Taking App",  # Optional: your app's name
            "Content-Type": "application/json"
        }
        
        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def test_connection(self) -> Dict[str, Any]:
        """Test the OpenRouter API connection."""
        try:
            self.logger.info("Testing OpenRouter API connection...")
            
            # Simple test request with minimal token usage
            test_payload = {
                "model": self.config.model,
                "messages": [
                    {"role": "user", "content": "Say 'Hello' if you can read this."}
                ],
                "max_tokens": 10,
                "temperature": 0.1
            }
            
            response = requests.post(
                self.chat_url,
                headers=self.headers,
                json=test_payload,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                self.logger.info("OpenRouter API connection successful")
                return {
                    "success": True,
                    "message": "Connection successful",
                    "model": self.config.model,
                    "response": result.get("choices", [{}])[0].get("message", {}).get("content", "")
                }
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                self.logger.error(f"OpenRouter API test failed: {error_msg}")
                return {
                    "success": False,
                    "error": error_msg
                }
        
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Network error testing OpenRouter API: {e}")
            return {
                "success": False,
                "error": f"Network error: {str(e)}"
            }
        except Exception as e:
            self.logger.error(f"Unexpected error testing OpenRouter API: {e}")
            return {
                "success": False,
                "error": f"Unexpected error: {str(e)}"
            }
    
    def get_available_models(self) -> List[Dict[str, Any]]:
        """Get list of available models from OpenRouter."""
        try:
            self.logger.info("Fetching available models from OpenRouter...")
            
            response = requests.get(
                self.models_url,
                headers=self.headers,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                models = result.get("data", [])
                self.logger.info(f"Retrieved {len(models)} available models")
                return models
            else:
                self.logger.error(f"Failed to fetch models: HTTP {response.status_code}")
                return []
        
        except Exception as e:
            self.logger.error(f"Error fetching available models: {e}")
            return []
    
    def process_transcript(self, transcript_text: str, subject: str = "") -> Dict[str, Any]:
        """Process transcript text into structured notes."""
        try:
            self.logger.info(f"Processing transcript for subject: {subject}")
            
            # Prepare the prompt
            system_prompt = self._get_system_prompt(subject)
            user_prompt = f"{self.config.pre_prompt}\n\n{transcript_text}" if self.config.pre_prompt else transcript_text
            
            # Prepare API request
            payload = {
                "model": self.config.model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": self.config.temperature,
                "max_tokens": self.config.max_tokens
            }
            
            # Make request with retries
            for attempt in range(self.config.max_retries):
                try:
                    response = requests.post(
                        self.chat_url,
                        headers=self.headers,
                        json=payload,
                        timeout=self.config.timeout
                    )
                    
                    if response.status_code == 200:
                        result = response.json()
                        
                        # Extract response content
                        choices = result.get("choices", [])
                        if not choices:
                            raise ValueError("No choices in API response")
                        
                        content = choices[0].get("message", {}).get("content", "")
                        if not content:
                            raise ValueError("Empty content in API response")
                        
                        # Extract token usage
                        usage = result.get("usage", {})
                        tokens_used = usage.get("total_tokens", 0)
                        
                        self.logger.info(f"Successfully processed transcript using {tokens_used} tokens")
                        
                        return {
                            "success": True,
                            "notes": content,
                            "tokens_used": tokens_used,
                            "model": self.config.model,
                            "usage": usage
                        }
                    
                    elif response.status_code == 429:  # Rate limited
                        wait_time = self.config.retry_delay * (2 ** attempt)
                        self.logger.warning(f"Rate limited, waiting {wait_time}s before retry {attempt + 1}")
                        time.sleep(wait_time)
                        continue
                    
                    else:
                        error_msg = f"HTTP {response.status_code}: {response.text}"
                        self.logger.error(f"API request failed: {error_msg}")
                        return {
                            "success": False,
                            "error": error_msg
                        }
                
                except requests.exceptions.Timeout:
                    if attempt < self.config.max_retries - 1:
                        self.logger.warning(f"Request timeout, retrying... (attempt {attempt + 1})")
                        time.sleep(self.config.retry_delay)
                        continue
                    else:
                        return {
                            "success": False,
                            "error": "Request timeout after all retries"
                        }
                
                except requests.exceptions.RequestException as e:
                    if attempt < self.config.max_retries - 1:
                        self.logger.warning(f"Network error, retrying... (attempt {attempt + 1}): {e}")
                        time.sleep(self.config.retry_delay)
                        continue
                    else:
                        return {
                            "success": False,
                            "error": f"Network error: {str(e)}"
                        }
            
            return {
                "success": False,
                "error": "All retry attempts failed"
            }
        
        except Exception as e:
            self.logger.error(f"Error processing transcript: {e}")
            return {
                "success": False,
                "error": f"Processing error: {str(e)}"
            }
    
    def process_transcript_file(self, transcript_path: str, output_path: str, subject: str = "") -> Dict[str, Any]:
        """Process transcript file and save notes to output file."""
        try:
            # Read transcript file
            transcript_file = Path(transcript_path)
            if not transcript_file.exists():
                return {
                    "success": False,
                    "error": f"Transcript file not found: {transcript_path}"
                }
            
            with open(transcript_file, 'r', encoding='utf-8') as f:
                transcript_text = f.read().strip()
            
            if not transcript_text:
                return {
                    "success": False,
                    "error": "Transcript file is empty"
                }
            
            # Process transcript
            result = self.process_transcript(transcript_text, subject)
            
            if result["success"]:
                # Save notes to file
                output_file = Path(output_path)
                output_file.parent.mkdir(parents=True, exist_ok=True)
                
                # Create formatted notes content
                notes_content = self._format_notes_output(
                    result["notes"],
                    subject,
                    transcript_file.name,
                    result.get("tokens_used", 0),
                    result.get("model", self.config.model)
                )
                
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(notes_content)
                
                self.logger.info(f"Notes saved to: {output_path}")
                
                return {
                    "success": True,
                    "output_path": output_path,
                    "tokens_used": result.get("tokens_used", 0),
                    "model": result.get("model", self.config.model)
                }
            else:
                return result
        
        except Exception as e:
            self.logger.error(f"Error processing transcript file: {e}")
            return {
                "success": False,
                "error": f"File processing error: {str(e)}"
            }
    
    def _get_system_prompt(self, subject: str = "") -> str:
        """Get system prompt for note processing."""
        base_prompt = """You are an expert note-taking assistant for students. Your task is to convert lecture transcripts into well-structured, comprehensive study notes.

Please transform the provided transcript into organized notes with these characteristics:
- Create clear headings and subheadings using markdown format
- Extract key concepts, definitions, and important facts
- Organize information logically and hierarchically
- Use bullet points and numbered lists where appropriate
- Highlight important terms and concepts with **bold** or *italics*
- Include examples and explanations provided in the lecture
- Maintain academic tone and accuracy
- Format for easy studying and review
- Add a brief summary at the end

Structure your response as:
# [Lecture Topic/Title]

## Key Concepts
[Main concepts covered]

## Detailed Notes
[Organized content with proper hierarchy]

## Important Definitions
[Key terms and their definitions]

## Examples
[Any examples mentioned in the lecture]

## Summary
[Brief overview of the main points]"""

        if subject:
            subject_addition = f"\n\nThis transcript is from a {subject} class, so focus on concepts and terminology relevant to that subject."
            return base_prompt + subject_addition
        
        return base_prompt
    
    def _format_notes_output(self, notes: str, subject: str, filename: str, tokens_used: int, model: str) -> str:
        """Format the notes output with metadata."""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        
        header = f"""---
Generated: {timestamp}
Source: {filename}
Subject: {subject}
Model: {model}
Tokens Used: {tokens_used}
---

"""
        
        return header + notes
    
    def estimate_tokens(self, text: str) -> int:
        """Rough estimation of token count for text."""
        # Very rough estimation: ~4 characters per token
        return len(text) // 4
    
    def get_model_info(self, model_id: str) -> Optional[Dict[str, Any]]:
        """Get information about a specific model."""
        models = self.get_available_models()
        for model in models:
            if model.get("id") == model_id:
                return model
        return None


# Example usage and testing
def main():
    """Example usage of the OpenRouter processor."""
    try:
        # Initialize processor
        config = NoteProcessingConfig(
            model="openai/gpt-4o-mini",
            temperature=0.3,
            max_tokens=2000
        )
        
        processor = OpenRouterProcessor(config=config)
        
        # Test connection
        print("Testing OpenRouter connection...")
        test_result = processor.test_connection()
        print(f"Connection test: {'✓' if test_result['success'] else '✗'}")
        
        if test_result['success']:
            print(f"Response: {test_result['response']}")
        else:
            print(f"Error: {test_result['error']}")
            return
        
        # Example transcript processing
        sample_transcript = """
        Good morning class. Today we're going to discuss photosynthesis, which is the process by which plants convert light energy into chemical energy. 
        
        The basic equation for photosynthesis is: 6CO2 + 6H2O + light energy → C6H12O6 + 6O2
        
        This process occurs in two main stages: the light reactions and the Calvin cycle. The light reactions occur in the thylakoids, while the Calvin cycle takes place in the stroma of chloroplasts.
        
        The light reactions capture energy from sunlight and use it to produce ATP and NADPH. These energy carriers are then used in the Calvin cycle to fix carbon dioxide into glucose.
        """
        
        print("\nProcessing sample transcript...")
        result = processor.process_transcript(sample_transcript, "Biology")
        
        if result['success']:
            print(f"✓ Processing successful! Used {result['tokens_used']} tokens")
            print("\nGenerated Notes:")
            print("-" * 50)
            print(result['notes'])
        else:
            print(f"✗ Processing failed: {result['error']}")
    
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()