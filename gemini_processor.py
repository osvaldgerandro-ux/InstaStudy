#!/usr/bin/env python3
"""
Gemini API Processor for Note Taking App
Handles communication with Google's Gemini API for AI-powered note processing.
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
class GeminiProcessingConfig:
    """Configuration for note processing with Gemini API."""
    model: str = "gemini-1.5-pro"
    temperature: float = 0.3
    max_tokens: int = 4000
    pre_prompt: str = ""
    timeout: int = 120
    max_retries: int = 3
    retry_delay: float = 1.0


class GeminiProcessor:
    """Handles note processing using Gemini API."""
    
    def __init__(self, api_key: str = None, config: GeminiProcessingConfig = None):
        self.config = config or GeminiProcessingConfig()
        
        # Get API key from parameter, environment, or file
        self.api_key = api_key
        if not self.api_key:
            self.api_key = os.getenv("GEMINI_API_KEY")
        if not self.api_key:
            api_key_file = Path("gemini_api_key.txt")
            if api_key_file.exists():
                try:
                    with open(api_key_file, 'r', encoding='utf-8') as f:
                        self.api_key = f.read().strip()
                except Exception as e:
                    logging.warning(f"Could not read API key from file: {e}")
        
        if not self.api_key:
            # We don't raise here to allow instantiation for testing/configuration
            # but methods requiring API key will fail
            pass
        
        # Setup logging
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)
    
    def test_connection(self) -> Dict[str, Any]:
        """Test the Gemini API connection."""
        try:
            if not self.api_key:
                return {
                    "success": False,
                    "error": "Gemini API key not found"
                }

            self.logger.info("Testing Gemini API connection...")
            
            # Simple test request
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.config.model}:generateContent?key={self.api_key}"
            
            payload = {
                "contents": [{
                    "parts": [{"text": "Say 'Hello' if you can read this."}]
                }],
                "generationConfig": {
                    "maxOutputTokens": 10,
                    "temperature": 0.1
                }
            }
            
            response = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                self.logger.info("Gemini API connection successful")
                
                # Extract text
                text = ""
                try:
                    text = result['candidates'][0]['content']['parts'][0]['text']
                except (KeyError, IndexError):
                    pass
                
                return {
                    "success": True,
                    "message": "Connection successful",
                    "model": self.config.model,
                    "response": text
                }
            else:
                error_msg = f"HTTP {response.status_code}: {response.text}"
                self.logger.error(f"Gemini API test failed: {error_msg}")
                return {
                    "success": False,
                    "error": error_msg
                }
        
        except requests.exceptions.RequestException as e:
            self.logger.error(f"Network error testing Gemini API: {e}")
            return {
                "success": False,
                "error": f"Network error: {str(e)}"
            }
        except Exception as e:
            self.logger.error(f"Unexpected error testing Gemini API: {e}")
            return {
                "success": False,
                "error": f"Unexpected error: {str(e)}"
            }
    
    def process_transcript(self, transcript_text: str, subject: str = "") -> Dict[str, Any]:
        """Process transcript text into structured notes."""
        try:
            if not self.api_key:
                return {
                    "success": False,
                    "error": "Gemini API key not found"
                }

            self.logger.info(f"Processing transcript for subject: {subject} with Gemini")
            
            # Prepare the prompt
            system_prompt = self._get_system_prompt(subject)
            full_prompt = f"{system_prompt}\n\n{self.config.pre_prompt}\n\nTranscript:\n{transcript_text}"
            
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{self.config.model}:generateContent?key={self.api_key}"
            
            payload = {
                "contents": [{
                    "parts": [{"text": full_prompt}]
                }],
                "generationConfig": {
                    "temperature": self.config.temperature,
                    "maxOutputTokens": self.config.max_tokens
                }
            }
            
            # Make request with retries
            for attempt in range(self.config.max_retries):
                try:
                    response = requests.post(
                        url,
                        headers={"Content-Type": "application/json"},
                        json=payload,
                        timeout=self.config.timeout
                    )
                    
                    if response.status_code == 200:
                        result = response.json()
                        
                        # Extract response content
                        content = ""
                        try:
                            content = result['candidates'][0]['content']['parts'][0]['text']
                        except (KeyError, IndexError):
                            raise ValueError("Unexpected response structure from Gemini API")
                        
                        if not content:
                            raise ValueError("Empty content in API response")
                        
                        # Extract token usage (if available, Gemini API format varies)
                        # usageMetadata: { promptTokenCount: 10, candidatesTokenCount: 100, totalTokenCount: 110 }
                        usage = result.get("usageMetadata", {})
                        tokens_used = usage.get("totalTokenCount", 0)
                        
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
                        # Don't retry for client errors unless 429
                        if 400 <= response.status_code < 500 and response.status_code != 429:
                             return {
                                "success": False,
                                "error": error_msg
                            }
                        
                        # Retry for 500s
                        if attempt < self.config.max_retries - 1:
                            time.sleep(self.config.retry_delay)
                            continue
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
