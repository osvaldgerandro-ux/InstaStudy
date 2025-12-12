#!/usr/bin/env python3
"""
Enhanced AssemblyAI Transcriber - A versatile transcription tool and API
Transcribe one or multiple local audio files using AssemblyAI.
- Reads API key from api_key.txt file automatically
- Supports automatic language detection
- Disables speaker diarization by default
- Provides both CLI and programmatic API interfaces
- Flexible output directory selection
- Multiple export formats (TXT, SRT, JSON)
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import List, Optional, Dict, Any, Union

# assemblyai SDK
import assemblyai as aai

# tkinter for file dialogs (GUI selection)
try:
    from tkinter import Tk
    from tkinter.filedialog import askopenfilenames, askdirectory
except Exception:
    # If tkinter is not present (headless environment) we'll fall back to CLI file arguments
    Tk = None
    askopenfilenames = None
    askdirectory = None


class TranscriptionConfig:
    """Configuration class for transcription settings."""
    
    def __init__(
        self,
        language_detection: bool = True,
        speaker_labels: bool = False,
        auto_highlights: bool = False,
        sentiment_analysis: bool = False,
        entity_detection: bool = False,
        punctuate: bool = True,
        format_text: bool = True,
        dual_channel: bool = False,
        webhook_url: Optional[str] = None
    ):
        self.language_detection = language_detection
        self.speaker_labels = speaker_labels
        self.auto_highlights = auto_highlights
        self.sentiment_analysis = sentiment_analysis
        self.entity_detection = entity_detection
        self.punctuate = punctuate
        self.format_text = format_text
        self.dual_channel = dual_channel
        self.webhook_url = webhook_url


class AssemblyAITranscriber:
    """
    Enhanced AssemblyAI Transcriber class providing programmatic API interface.
    """
    
    def __init__(self, api_key: Optional[str] = None, config: Optional[TranscriptionConfig] = None):
        """
        Initialize the transcriber.
        
        Args:
            api_key: AssemblyAI API key. If None, will try to load from api_key.txt
            config: TranscriptionConfig object with transcription settings
        """
        self.api_key = self._get_api_key(api_key)
        self.config = config or TranscriptionConfig()
        
        # Set SDK API key
        aai.settings.api_key = self.api_key
        self.transcriber = aai.Transcriber()
    
    def _get_api_key(self, provided_key: Optional[str] = None) -> str:
        """Get API key from various sources in order of preference."""
        if provided_key:
            return provided_key
        
        # Try to read from api_key.txt in the same directory as the script
        script_dir = Path(__file__).parent
        api_key_file = script_dir / "api_key.txt"
        
        if api_key_file.exists():
            try:
                with open(api_key_file, 'r', encoding='utf-8') as f:
                    key = f.read().strip()
                    if key:
                        return key
            except Exception as e:
                print(f"Warning: Could not read api_key.txt: {e}")
        
        # Try environment variable
        env_key = os.getenv("ASSEMBLYAI_API_KEY")
        if env_key:
            return env_key
        
        raise ValueError("API key not found. Please provide it via parameter, api_key.txt file, or ASSEMBLYAI_API_KEY env var.")
    
    def _create_transcript_config(self, language_code: Optional[str] = None) -> aai.TranscriptionConfig:
        """Create AssemblyAI TranscriptionConfig from our config."""
        transcript_config = aai.TranscriptionConfig(
            language_detection=self.config.language_detection if not language_code else False,
            language_code=language_code,
            speaker_labels=self.config.speaker_labels,
            auto_highlights=self.config.auto_highlights,
            sentiment_analysis=self.config.sentiment_analysis,
            entity_detection=self.config.entity_detection,
            punctuate=self.config.punctuate,
            format_text=self.config.format_text,
            dual_channel=self.config.dual_channel,
            webhook_url=self.config.webhook_url
        )
        return transcript_config
    
    def transcribe_file(
        self, 
        filepath: Union[str, Path], 
        output_dir: Optional[Union[str, Path]] = None,
        language_code: Optional[str] = None,
        save_txt: bool = True,
        save_srt: bool = False,
        save_json: bool = False,
        custom_filename: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Transcribe a single audio file.
        
        Args:
            filepath: Path to the audio file
            output_dir: Directory to save output files. If None, saves next to source file
            language_code: Specific language code (overrides auto-detection)
            save_txt: Save plain text transcript
            save_srt: Save SRT subtitle file
            save_json: Save full JSON response
            custom_filename: Custom base filename for outputs
        
        Returns:
            Dict containing transcription results and metadata
        """
        filepath = Path(filepath)
        if not filepath.exists():
            raise FileNotFoundError(f"Audio file not found: {filepath}")
        
        # Determine output directory
        if output_dir:
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
        else:
            output_dir = filepath.parent
        
        # Determine base filename
        if custom_filename:
            base_name = custom_filename
        else:
            base_name = filepath.stem
        
        print(f"\nTranscribing: {filepath.name}")
        
        try:
            # Create transcription config
            transcript_config = self._create_transcript_config(language_code)
            
            # Transcribe
            transcript = self.transcriber.transcribe(str(filepath), config=transcript_config)
            
            if transcript.error:
                raise Exception(f"Transcription error: {transcript.error}")
            
            result = {
                'success': True,
                'filepath': str(filepath),
                'transcript_id': transcript.id,
                'text': transcript.text,
                'language_detected': getattr(transcript, 'language_code', None),
                'confidence': getattr(transcript, 'confidence', None),
                'audio_duration': getattr(transcript, 'audio_duration', None),
                'output_files': []
            }
            
            # Save outputs
            if save_txt and transcript.text:
                txt_file = output_dir / f"{base_name}.txt"
                with open(txt_file, 'w', encoding='utf-8') as f:
                    f.write(transcript.text)
                result['output_files'].append(str(txt_file))
                print(f"  ✓ Saved transcript to {txt_file}")
            
            if save_srt:
                try:
                    srt_content = transcript.export_subtitles_srt()
                    srt_file = output_dir / f"{base_name}.srt"
                    with open(srt_file, 'w', encoding='utf-8') as f:
                        f.write(srt_content)
                    result['output_files'].append(str(srt_file))
                    print(f"  ✓ Saved SRT to {srt_file}")
                except Exception as e:
                    print(f"  ! Could not export SRT: {e}")
            
            if save_json:
                json_file = output_dir / f"{base_name}.json"
                json_data = {
                    'id': transcript.id,
                    'text': transcript.text,
                    'status': transcript.status,
                    'language_code': getattr(transcript, 'language_code', None),
                    'confidence': getattr(transcript, 'confidence', None),
                    'audio_duration': getattr(transcript, 'audio_duration', None),
                }
                
                # Add words if available (handle different SDK versions)
                if hasattr(transcript, 'words') and transcript.words:
                    try:
                        # Try model_dump() first (newer SDK versions)
                        json_data['words'] = [word.model_dump() for word in transcript.words]
                    except AttributeError:
                        # Fallback to dict conversion for older SDK versions
                        json_data['words'] = []
                        for word in transcript.words:
                            word_dict = {
                                'text': getattr(word, 'text', ''),
                                'start': getattr(word, 'start', 0),
                                'end': getattr(word, 'end', 0),
                                'confidence': getattr(word, 'confidence', 0.0)
                            }
                            json_data['words'].append(word_dict)
                
                # Add optional features if enabled
                if self.config.speaker_labels and hasattr(transcript, 'utterances') and transcript.utterances:
                    try:
                        json_data['utterances'] = [utt.model_dump() for utt in transcript.utterances]
                    except AttributeError:
                        # Fallback for utterances
                        json_data['utterances'] = []
                        for utt in transcript.utterances:
                            utt_dict = {
                                'text': getattr(utt, 'text', ''),
                                'start': getattr(utt, 'start', 0),
                                'end': getattr(utt, 'end', 0),
                                'confidence': getattr(utt, 'confidence', 0.0),
                                'speaker': getattr(utt, 'speaker', 'Unknown')
                            }
                            json_data['utterances'].append(utt_dict)
                
                if self.config.sentiment_analysis and hasattr(transcript, 'sentiment_analysis_results'):
                    json_data['sentiment_analysis'] = transcript.sentiment_analysis_results
                
                if self.config.entity_detection and hasattr(transcript, 'entities') and transcript.entities:
                    try:
                        json_data['entities'] = [entity.model_dump() for entity in transcript.entities]
                    except AttributeError:
                        # Fallback for entities
                        json_data['entities'] = []
                        for entity in transcript.entities:
                            entity_dict = {
                                'text': getattr(entity, 'text', ''),
                                'entity_type': getattr(entity, 'entity_type', ''),
                                'start': getattr(entity, 'start', 0),
                                'end': getattr(entity, 'end', 0)
                            }
                            json_data['entities'].append(entity_dict)
                
                if self.config.auto_highlights and hasattr(transcript, 'auto_highlights'):
                    json_data['auto_highlights'] = transcript.auto_highlights
                
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, indent=2, ensure_ascii=False)
                result['output_files'].append(str(json_file))
                print(f"  ✓ Saved JSON data to {json_file}")
            
            return result
            
        except Exception as e:
            print(f"  ✖ Error transcribing {filepath.name}: {e}")
            return {
                'success': False,
                'filepath': str(filepath),
                'error': str(e),
                'output_files': []
            }
    
    def transcribe_files(
        self,
        filepaths: List[Union[str, Path]],
        output_dir: Optional[Union[str, Path]] = None,
        language_code: Optional[str] = None,
        save_txt: bool = True,
        save_srt: bool = False,
        save_json: bool = False
    ) -> List[Dict[str, Any]]:
        """
        Transcribe multiple audio files.
        
        Args:
            filepaths: List of paths to audio files
            output_dir: Directory to save output files
            language_code: Specific language code (overrides auto-detection)
            save_txt: Save plain text transcripts
            save_srt: Save SRT subtitle files
            save_json: Save full JSON responses
        
        Returns:
            List of transcription results
        """
        results = []
        print(f"Will transcribe {len(filepaths)} file(s).")
        
        for filepath in filepaths:
            result = self.transcribe_file(
                filepath=filepath,
                output_dir=output_dir,
                language_code=language_code,
                save_txt=save_txt,
                save_srt=save_srt,
                save_json=save_json
            )
            results.append(result)
        
        return results


# GUI Helper Functions
def pick_files_with_tkinter(multiple=True, title="Select audio file(s)"):
    """Pick files using tkinter file dialog."""
    if Tk is None or askopenfilenames is None:
        return []
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    filetypes = [
        ("Audio files", ("*.wav", "*.mp3", "*.m4a", "*.flac", "*.aac", "*.ogg", "*.mp4", "*.webm")),
        ("All files", "*.*"),
    ]
    paths = askopenfilenames(title=title, filetypes=filetypes)
    root.destroy()
    return list(paths)


def pick_directory_with_tkinter(title="Select output directory"):
    """Pick directory using tkinter directory dialog."""
    if Tk is None or askdirectory is None:
        return ""
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    directory = askdirectory(title=title)
    root.destroy()
    return directory


# CLI Interface
def main():
    parser = argparse.ArgumentParser(
        description="Enhanced AssemblyAI Transcriber - Transcribe audio files with advanced features.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # GUI mode (default)
  python transcriber.py
  
  # CLI mode with specific files
  python transcriber.py --no-gui --files audio1.mp3 audio2.wav
  
  # Specify output directory and formats
  python transcriber.py --files audio.mp3 --output-dir ./transcripts --srt --json
  
  # Use specific language
  python transcriber.py --files audio.mp3 --language en
  
  # Enable advanced features
  python transcriber.py --files audio.mp3 --speaker-labels --sentiment-analysis
        """
    )
    
    # Basic options
    parser.add_argument("--api-key", help="AssemblyAI API key (or place in api_key.txt file).")
    parser.add_argument("--no-gui", action="store_true", help="Don't open file dialog (use --files instead).")
    parser.add_argument("--files", nargs="+", help="Paths to audio files.")
    parser.add_argument("--output-dir", help="Output directory for transcription files.")
    parser.add_argument("--language", help="Language code (e.g., 'en', 'es', 'fr'). If not specified, auto-detection is used.")
    
    # Output formats
    parser.add_argument("--srt", action="store_true", help="Save SRT subtitle files.")
    parser.add_argument("--json", action="store_true", help="Save detailed JSON files with metadata.")
    parser.add_argument("--no-txt", action="store_true", help="Don't save plain text files.")
    
    # Advanced features
    parser.add_argument("--speaker-labels", action="store_true", help="Enable speaker diarization/labeling.")
    parser.add_argument("--sentiment-analysis", action="store_true", help="Enable sentiment analysis.")
    parser.add_argument("--entity-detection", action="store_true", help="Enable entity detection.")
    parser.add_argument("--auto-highlights", action="store_true", help="Enable automatic highlights.")
    parser.add_argument("--dual-channel", action="store_true", help="Process dual-channel audio separately.")
    
    args = parser.parse_args()
    
    try:
        # Create transcription config
        config = TranscriptionConfig(
            language_detection=not args.language,  # Disable if specific language provided
            speaker_labels=args.speaker_labels,
            sentiment_analysis=args.sentiment_analysis,
            entity_detection=args.entity_detection,
            auto_highlights=args.auto_highlights,
            dual_channel=args.dual_channel
        )
        
        # Initialize transcriber
        transcriber = AssemblyAITranscriber(api_key=args.api_key, config=config)
        
        # Determine file list
        files = []
        if not args.no_gui:
            files = pick_files_with_tkinter()
            # Also allow GUI selection of output directory
            if files and not args.output_dir:
                output_dir = pick_directory_with_tkinter()
                if output_dir:
                    args.output_dir = output_dir
        
        # Fallback to --files if GUI returned nothing or --no-gui
        if not files and args.files:
            files = args.files
        
        if not files:
            print("No files selected. Use --files FILE1 FILE2 ... or run without --no-gui to pick files with a dialog.")
            sys.exit(0)
        
        # Transcribe files
        results = transcriber.transcribe_files(
            filepaths=files,
            output_dir=args.output_dir,
            language_code=args.language,
            save_txt=not args.no_txt,
            save_srt=args.srt,
            save_json=args.json
        )
        
        # Print summary
        successful = sum(1 for r in results if r['success'])
        print(f"\n=== Summary ===")
        print(f"Total files: {len(results)}")
        print(f"Successful: {successful}")
        print(f"Failed: {len(results) - successful}")
        
        # Exit with error code if any failures
        if successful < len(results):
            sys.exit(1)
            
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(2)


if __name__ == "__main__":
    main()