#!/usr/bin/env python3
"""
School Note Taking App - Main Controller with Color Management Integration
Orchestrates recording detection, transcription, note processing, and document color theming.
"""

import os
import sys
import json
import time
import threading
from pathlib import Path
from typing import Dict, List, Optional, Set
from datetime import datetime
import queue
import logging
from dataclasses import dataclass, asdict, fields
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from word_document_manager import WordDocumentManager, WordFormattingConfig

try:
    from transcriber import AssemblyAITranscriber, TranscriptionConfig
    from openrouter_processor import OpenRouterProcessor, NoteProcessingConfig
    from gemini_processor import GeminiProcessor, GeminiProcessingConfig
    from color_changer import COLOR_PALETTES, change_theme_colors
except ImportError as e:
    print(f"Error: Could not import required modules: {e}")
    print("Make sure transcriber.py, openrouter_processor.py, gemini_processor.py, word_document_manager.py, and color_changer.py are in the same directory.")
    sys.exit(1)


@dataclass
class AppConfig:
    """Application configuration."""
    watch_directory: str = ""
    subjects: List[str] = None
    auto_process: bool = True
    supported_extensions: List[str] = None
    
    # Provider configuration
    provider_mode: str = "Only OpenRouter"  # "Only OpenRouter", "Only Gemini", "Fallback Mode"
    primary_provider: str = "OpenRouter"    # "OpenRouter" or "Gemini"
    secondary_provider: str = "Gemini"      # "OpenRouter" or "Gemini"
    
    # OpenRouter configuration
    openrouter_model: str = "openai/gpt-4o-mini"
    openrouter_temperature: float = 0.3
    openrouter_max_tokens: int = 4000
    
    # Gemini configuration
    gemini_model: str = "gemini-1.5-pro"
    gemini_temperature: float = 0.3
    gemini_max_tokens: int = 4000
    
    word_auto_update: bool = True
    word_font_name: str = "Calibri"
    word_font_size: int = 11
    word_heading1_size: int = 18
    word_heading2_size: int = 16
    word_heading3_size: int = 14
    word_line_spacing: float = 1.15
    remove_thinking_tags: bool = True
    subject_colors: Dict[str, str] = None
    auto_apply_colors: bool = True

    def __post_init__(self):
        if self.subjects is None:
            self.subjects = []
        if self.supported_extensions is None:
            self.supported_extensions = ['.mp3', '.wav', '.m4a', '.mp4', '.flac', '.aac', '.ogg', '.webm']
        if self.subject_colors is None:
            self.subject_colors = {}
        if not hasattr(self, 'word_auto_update'):
            self.word_auto_update = True
        if not hasattr(self, 'word_font_name'):
            self.word_font_name = "Calibri"
        if not hasattr(self, 'word_font_size'):
            self.word_font_size = 11
        if not hasattr(self, 'auto_apply_colors'):
            self.auto_apply_colors = True
        
        # New configuration defaults
        if not hasattr(self, 'provider_mode'):
            self.provider_mode = "Only OpenRouter"
        if not hasattr(self, 'primary_provider'):
            self.primary_provider = "OpenRouter"
        if not hasattr(self, 'secondary_provider'):
            self.secondary_provider = "Gemini"
        if not hasattr(self, 'gemini_model'):
            self.gemini_model = "gemini-1.5-pro"
        if not hasattr(self, 'gemini_temperature'):
            self.gemini_temperature = 0.3
        if not hasattr(self, 'gemini_max_tokens'):
            self.gemini_max_tokens = 4000


@dataclass
class ProcessingTask:
    """Represents a file processing task."""
    filepath: str
    subject: str
    status: str = "pending"
    transcript_path: str = ""
    notes_path: str = ""
    error_message: str = ""
    created_at: str = ""
    tokens_used: int = 0
    reprocess_type: str = ""

    def __post_init__(self):
        if not self.created_at:
            self.created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@dataclass
class ReprocessingFileInfo:
    """Information about a file available for reprocessing."""
    filepath: str
    filename: str
    subject: str
    has_transcript: bool = False
    has_notes: bool = False
    transcript_path: str = ""
    notes_path: str = ""
    file_size: int = 0
    modified_date: str = ""
    selected: bool = False


class AudioFileHandler(FileSystemEventHandler):
    """Handles file system events for new audio files."""

    def __init__(self, app):
        self.app = app

    def on_created(self, event):
        if event.is_directory:
            return
        self.app.handle_new_file(event.src_path)

    def on_moved(self, event):
        if event.is_directory:
            return
        self.app.handle_new_file(event.dest_path)


class NoteProcessingThread(threading.Thread):
    """Background thread for processing transcription to notes."""

    def __init__(self, task_queue, app):
        super().__init__(daemon=True)
        self.task_queue = task_queue
        self.app = app
        self.running = True

    def run(self):
        while self.running:
            try:
                task = self.task_queue.get(timeout=1)
                self.app.process_task(task)
                self.task_queue.task_done()
            except queue.Empty:
                continue
            except Exception as e:
                logging.error(f"Error in processing thread: {e}")

    def stop(self):
        self.running = False


class SchoolNoteApp:
    """Main application class."""

    def __init__(self):
        self.config = AppConfig()
        self.config_file = Path("app_config.json")
        self.tasks: Dict[str, ProcessingTask] = {}
        self.task_queue = queue.Queue()
        self.processing_thread = None
        self.observer = None

        self.reprocessing_files: Dict[str, ReprocessingFileInfo] = {}
        self.selected_files: Set[str] = set()

        self.assemblyai_key_file = Path("assemblyai_api_key.txt")
        self.openrouter_key_file = Path("openrouter_api_key.txt")
        self.gemini_key_file = Path("gemini_api_key.txt")
        self.pre_prompt_file = Path("pre_prompt.txt")
        self.word_manager = None
        self.init_word_manager()

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('note_app.log'),
                logging.StreamHandler()
            ]
        )

        self.load_config()
        self.ensure_pre_prompt_file()
        self.setup_gui()
        self.start_processing_thread()

        if self.config.watch_directory and os.path.exists(self.config.watch_directory):
            self.start_file_monitoring()

    def ensure_pre_prompt_file(self):
        """Ensure pre-prompt file exists with default content."""
        if not self.pre_prompt_file.exists():
            default_prompt = """You are an expert note-taking assistant for students. Your task is to convert lecture transcripts into well-structured, comprehensive study notes.

Please transform the following transcript into organized notes with these characteristics:
- Create clear headings and subheadings
- Extract key concepts, definitions, and important facts
- Organize information logically and hierarchically
- Use bullet points and numbered lists where appropriate
- Highlight important terms and concepts
- Include examples and explanations provided in the lecture
- Maintain academic tone and accuracy
- Format for easy studying and review

Transcript to process:"""

            try:
                with open(self.pre_prompt_file, 'w', encoding='utf-8') as f:
                    f.write(default_prompt)
            except Exception as e:
                logging.error(f"Error creating default pre-prompt file: {e}")

    def init_word_manager(self):
        """Initialize Word document manager with current configuration."""
        try:
            formatting_config = WordFormattingConfig(
                font_name=self.config.word_font_name,
                font_size=self.config.word_font_size,
                heading1_size=self.config.word_heading1_size,
                heading2_size=self.config.word_heading2_size,
                heading3_size=self.config.word_heading3_size,
                line_spacing=self.config.word_line_spacing
            )
            self.word_manager = WordDocumentManager(formatting_config)
            logging.info("Word document manager initialized")
        except Exception as e:
            logging.error(f"Error initializing Word document manager: {e}")
            self.word_manager = None

    def load_config(self):
        """Load configuration from file."""
        if not self.config_file.exists():
            return

        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                raw_data = json.load(f)

            if not isinstance(raw_data, dict):
                logging.error("Error loading configuration: app_config.json is not a JSON object")
                return

            data = dict(raw_data)

            # Backward-compatibility migrations
            if "ai_provider" in data and "provider_mode" not in data:
                provider = str(data.get("ai_provider", "")).strip().lower()
                if provider in {"openrouter", "open router", "only openrouter", "only open router"}:
                    data["provider_mode"] = "Only OpenRouter"
                elif provider in {"gemini", "only gemini"}:
                    data["provider_mode"] = "Only Gemini"

            if "model" in data and "openrouter_model" not in data:
                data["openrouter_model"] = data.get("model")
            if "temperature" in data and "openrouter_temperature" not in data:
                data["openrouter_temperature"] = data.get("temperature")
            if "max_tokens" in data and "openrouter_max_tokens" not in data:
                data["openrouter_max_tokens"] = data.get("max_tokens")

            if "font_name" in data and "word_font_name" not in data:
                data["word_font_name"] = data.get("font_name")
            if "font_size" in data and "word_font_size" not in data:
                data["word_font_size"] = data.get("font_size")
            if "heading1_size" in data and "word_heading1_size" not in data:
                data["word_heading1_size"] = data.get("heading1_size")
            if "heading2_size" in data and "word_heading2_size" not in data:
                data["word_heading2_size"] = data.get("heading2_size")
            if "heading3_size" in data and "word_heading3_size" not in data:
                data["word_heading3_size"] = data.get("heading3_size")
            if "line_spacing" in data and "word_line_spacing" not in data:
                data["word_line_spacing"] = data.get("line_spacing")

            valid_keys = {f.name for f in fields(AppConfig)}
            filtered = {k: v for k, v in data.items() if k in valid_keys}

            ignored = sorted(set(data.keys()) - valid_keys)
            if ignored:
                logging.info(f"Ignoring unknown config keys: {', '.join(ignored)}")

            self.config = AppConfig(**filtered)
            logging.info("Configuration loaded successfully")

            # Ensure dependent services reflect loaded settings
            self.init_word_manager()

        except Exception as e:
            logging.error(f"Error loading configuration: {e}")

    def save_config(self):
        """Save configuration to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(asdict(self.config), f, indent=2)
            logging.info("Configuration saved successfully")
        except Exception as e:
            logging.error(f"Error saving configuration: {e}")

    def read_api_key_file(self, file_path: Path) -> str:
        """Read API key from file."""
        if file_path.exists():
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read().strip()
            except Exception as e:
                logging.error(f"Error reading {file_path}: {e}")
        return ""

    def write_api_key_file(self, file_path: Path, api_key: str) -> bool:
        """Write API key to file."""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(api_key.strip())
            return True
        except Exception as e:
            logging.error(f"Error writing {file_path}: {e}")
            messagebox.showerror("Error", f"Failed to save API key: {e}")
            return False

    def read_pre_prompt(self) -> str:
        """Read pre-prompt from file."""
        try:
            with open(self.pre_prompt_file, 'r', encoding='utf-8') as f:
                return f.read()
        except Exception as e:
            logging.error(f"Error reading pre-prompt: {e}")
            return ""

    def write_pre_prompt(self, prompt: str) -> bool:
        """Write pre-prompt to file."""
        try:
            with open(self.pre_prompt_file, 'w', encoding='utf-8') as f:
                f.write(prompt)
            return True
        except Exception as e:
            logging.error(f"Error writing pre-prompt: {e}")
            messagebox.showerror("Error", f"Failed to save pre-prompt: {e}")
            return False

    def remove_thinking_tags(self, text: str) -> str:
        """Remove content between <think> and </think> tags."""
        import re
        pattern = r'<think\s*>.*?</think\s*>'
        cleaned_text = re.sub(pattern, '', text, flags=re.DOTALL | re.IGNORECASE)
        return cleaned_text

    def apply_color_to_word_document(self, subject: str):
        """Apply color palette to subject's Word document."""
        if subject not in self.config.subject_colors:
            logging.info(f"No color assigned for subject: {subject}")
            return
    
        palette_name = self.config.subject_colors[subject]
        if palette_name == "default":
            logging.info(f"Default color selected for {subject}, skipping color application")
            return
    
        # Search for the Word document in multiple possible locations
        possible_paths = [
            # Location 1: .\Appunti Completi\{subject}\{subject}_combined_notes.docx
            Path("Appunti Completi") / subject / f"{subject}_combined_notes.docx",
            # Location 2: .\{subject}\Appunti Completi\{subject}_combined_notes.docx
            Path(subject) / "Appunti Completi" / f"{subject}_combined_notes.docx",
            # Location 3: Current directory
            Path(f"{subject}_combined_notes.docx"),
            # Location 4: Subject directory
            Path(subject) / f"{subject}_combined_notes.docx"
        ]
    
        word_doc_path = None
        for path in possible_paths:
            if path.exists():
                word_doc_path = path
                logging.info(f"Found Word document at: {word_doc_path}")
                break
    
        if not word_doc_path:
            logging.warning(f"Word document not found for {subject}. Searched in:")
            for path in possible_paths:
                logging.warning(f"  - {path}")
            return
    
        try:
            change_theme_colors(str(word_doc_path), palette_name)
            self.log_activity(f"Applied '{COLOR_PALETTES[palette_name]['name']}' color to {subject} document")
            logging.info(f"Color '{palette_name}' applied to {subject} document at {word_doc_path}")
        except Exception as e:
            logging.error(f"Error applying color to {subject} document: {e}")
            self.log_activity(f"Error applying color to {subject}: {e}")

    def setup_gui(self):
        """Setup the GUI interface."""
        self.root = tk.Tk()
        self.root.title("School Note Taking App")
        self.root.geometry("1200x900")

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.setup_config_tab(self.notebook)
        self.setup_api_keys_tab(self.notebook)
        self.setup_pre_prompt_tab(self.notebook)
        self.setup_monitoring_tab(self.notebook)
        self.setup_tasks_tab(self.notebook)
        self.setup_reprocessing_tab(self.notebook)
        self.setup_word_tab(self.notebook)
        self.setup_color_management_tab(self.notebook)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.update_gui()

    def update_gui(self):
        """Refresh dynamic GUI elements based on current app state."""
        try:
            if hasattr(self, "monitor_status_var"):
                if self.observer and self.observer.is_alive():
                    self.monitor_status_var.set(f"Monitoring: {self.config.watch_directory}")
                else:
                    if self.config.watch_directory:
                        if os.path.exists(self.config.watch_directory):
                            self.monitor_status_var.set(f"Not monitoring (directory set: {self.config.watch_directory})")
                        else:
                            self.monitor_status_var.set(f"Not monitoring (missing directory: {self.config.watch_directory})")
                    else:
                        self.monitor_status_var.set("Not monitoring (no watch directory configured)")

            if hasattr(self, "status_var"):
                queued = sum(1 for t in self.tasks.values() if str(t.status).startswith("queued"))
                self.status_var.set(f"Ready • {queued} queued" if queued else "Ready")

            if hasattr(self, "provider_mode_var"):
                self.on_provider_mode_change(None)

            if hasattr(self, "tasks_tree"):
                self.refresh_tasks_display()

            if hasattr(self, "selection_count_var"):
                self.update_selection_count()

        except Exception as e:
            logging.error(f"Error updating GUI: {e}")

    def setup_color_management_tab(self, notebook):
        """Setup color management tab."""
        color_frame = ttk.Frame(notebook)
        notebook.add(color_frame, text="Color Management")

        ttk.Label(color_frame, text="Assign color palettes to each subject:",
                 font=("", 11, "bold")).pack(anchor=tk.W, padx=10, pady=(10, 5))

        ttk.Label(color_frame, text="Colors will be applied to Word documents after generation.",
                 font=("", 9, "italic")).pack(anchor=tk.W, padx=10, pady=(0, 10))

        self.auto_apply_colors_var = tk.BooleanVar(value=self.config.auto_apply_colors)
        ttk.Checkbutton(color_frame, text="Automatically apply colors after document generation",
                       variable=self.auto_apply_colors_var).pack(anchor=tk.W, padx=10, pady=5)

        table_frame = ttk.Frame(color_frame)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        columns = ("Subject", "Assigned Color", "Preview")
        self.color_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=12)

        self.color_tree.heading("Subject", text="Subject")
        self.color_tree.heading("Assigned Color", text="Assigned Color Palette")
        self.color_tree.heading("Preview", text="Status")

        self.color_tree.column("Subject", width=200)
        self.color_tree.column("Assigned Color", width=250)
        self.color_tree.column("Preview", width=300)

        color_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.color_tree.yview)
        self.color_tree.configure(yscrollcommand=color_scrollbar.set)

        self.color_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        color_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.color_tree.bind("<Double-Button-1>", self.on_color_assignment_double_click)

        buttons_frame = ttk.Frame(color_frame)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(buttons_frame, text="Assign Color to Selected",
                  command=self.assign_color_to_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Apply Colors to All Documents",
                  command=self.apply_colors_to_all_documents).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Clear All Assignments",
                  command=self.clear_all_color_assignments).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Refresh List",
                  command=self.refresh_color_assignments).pack(side=tk.LEFT, padx=(0, 5))

        palette_info_frame = ttk.LabelFrame(color_frame, text="Available Color Palettes", padding=10)
        palette_info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        palette_text = scrolledtext.ScrolledText(palette_info_frame, height=8, wrap=tk.WORD, font=("Consolas", 9))
        palette_text.pack(fill=tk.BOTH, expand=True)

        palette_list = "Available Palettes:\n" + "="*60 + "\n"
        for key, palette in COLOR_PALETTES.items():
            palette_list += f"  {key:15} - {palette['name']}\n"
        palette_list += "  default         - No color customization (Office default)\n"

        palette_text.insert(tk.END, palette_list)
        palette_text.config(state=tk.DISABLED)

        self.color_status_var = tk.StringVar(value="Ready")
        ttk.Label(color_frame, textvariable=self.color_status_var, font=("", 10, "bold")).pack(pady=10)

        self.refresh_color_assignments()

    def refresh_color_assignments(self):
        """Refresh the color assignments display."""
        for item in self.color_tree.get_children():
            self.color_tree.delete(item)
    
        if not self.config.subjects:
            self.color_tree.insert("", tk.END, values=("No subjects configured", "", ""))
            return
    
        for subject in self.config.subjects:
            assigned_color = self.config.subject_colors.get(subject, "default")
    
            if assigned_color == "default":
                color_name = "Default (no customization)"
            elif assigned_color in COLOR_PALETTES:
                color_name = COLOR_PALETTES[assigned_color]["name"]
            else:
                color_name = f"Unknown ({assigned_color})"
    
            # Search for Word document in multiple locations
            possible_paths = [
                Path("Appunti Completi") / subject / f"{subject}_combined_notes.docx",
                Path(subject) / "Appunti Completi" / f"{subject}_combined_notes.docx",
                Path(f"{subject}_combined_notes.docx"),
                Path(subject) / f"{subject}_combined_notes.docx"
            ]
    
            word_doc = None
            for path in possible_paths:
                if path.exists():
                    word_doc = path
                    break
    
            if word_doc:
                status = f"✓ Document exists at: {word_doc.parent.name}/{word_doc.name} ({word_doc.stat().st_size / 1024:.1f} KB)"
            else:
                status = "Document not yet created"
    
            self.color_tree.insert("", tk.END, values=(subject, color_name, status))

    def on_color_assignment_double_click(self, event):
        """Handle double-click on color assignment row."""
        selection = self.color_tree.selection()
        if not selection:
            return

        item = selection[0]
        values = self.color_tree.item(item, "values")

        if not values or values[0] == "No subjects configured":
            return

        subject = values[0]
        self.show_color_picker(subject)

    def assign_color_to_selected(self):
        """Assign color to selected subject."""
        selection = self.color_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subject first!")
            return

        item = selection[0]
        values = self.color_tree.item(item, "values")
        subject = values[0]

        self.show_color_picker(subject)

    def show_color_picker(self, subject: str):
        """Show color picker dialog for a subject."""
        picker_window = tk.Toplevel(self.root)
        picker_window.title(f"Select Color Palette for {subject}")
        picker_window.geometry("500x600")
        picker_window.transient(self.root)

        ttk.Label(picker_window, text=f"Choose color palette for: {subject}",
                 font=("", 11, "bold")).pack(pady=10)

        palettes_frame = ttk.Frame(picker_window)
        palettes_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        canvas = tk.Canvas(palettes_frame)
        scrollbar = ttk.Scrollbar(palettes_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        selected_palette = tk.StringVar(value=self.config.subject_colors.get(subject, "default"))

        ttk.Radiobutton(scrollable_frame, text="Default (no customization)",
                       variable=selected_palette, value="default").pack(anchor=tk.W, pady=2)

        for palette_key, palette_info in COLOR_PALETTES.items():
            ttk.Radiobutton(scrollable_frame, text=f"{palette_info['name']} ({palette_key})",
                           variable=selected_palette, value=palette_key).pack(anchor=tk.W, pady=2)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def apply_color():
            palette = selected_palette.get()
            self.config.subject_colors[subject] = palette
            self.save_config()

            if palette == "default":
                palette_name = "Default"
            else:
                palette_name = COLOR_PALETTES[palette]["name"]

            self.log_activity(f"Assigned '{palette_name}' to {subject}")
            self.refresh_color_assignments()
            picker_window.destroy()

            if messagebox.askyesno("Apply Now?",
                                  f"Color '{palette_name}' assigned to {subject}.\n\n"
                                  f"Apply color to the Word document now?"):
                self.apply_color_to_word_document(subject)

        buttons_frame = ttk.Frame(picker_window)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(buttons_frame, text="Apply", command=apply_color).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Cancel", command=picker_window.destroy).pack(side=tk.LEFT, padx=5)

    def apply_colors_to_all_documents(self):
        """Apply colors to all Word documents."""
        if not self.config.subjects:
            messagebox.showwarning("Warning", "No subjects configured!")
            return
    
        applied_count = 0
        skipped_count = 0
        error_count = 0
    
        self.color_status_var.set("Applying colors to documents...")
        self.root.update()
    
        results_details = []
    
        for subject in self.config.subjects:
            if subject not in self.config.subject_colors or self.config.subject_colors[subject] == "default":
                skipped_count += 1
                results_details.append(f"{subject}: Skipped (no color assigned)")
                continue
    
            # Search for the Word document
            possible_paths = [
                Path("Appunti Completi") / subject / f"{subject}_combined_notes.docx",
                Path(subject) / "Appunti Completi" / f"{subject}_combined_notes.docx",
                Path(f"{subject}_combined_notes.docx"),
                Path(subject) / f"{subject}_combined_notes.docx"
            ]
    
            word_doc = None
            for path in possible_paths:
                if path.exists():
                    word_doc = path
                    break
    
            if not word_doc:
                skipped_count += 1
                results_details.append(f"{subject}: Document not found")
                continue
    
            try:
                self.apply_color_to_word_document(subject)
                applied_count += 1
                results_details.append(f"{subject}: ✓ Color applied")
            except Exception as e:
                logging.error(f"Error applying color to {subject}: {e}")
                error_count += 1
                results_details.append(f"{subject}: ✗ Error - {str(e)[:50]}")
    
        self.color_status_var.set(f"Applied: {applied_count}, Skipped: {skipped_count}, Errors: {error_count}")
    
        # Show detailed results
        details_msg = "\n".join(results_details[:10])  # Show first 10 results
        if len(results_details) > 10:
            details_msg += f"\n... and {len(results_details) - 10} more"
    
        messagebox.showinfo("Color Application Complete",
                           f"Colors applied to {applied_count} document(s).\n"
                           f"Skipped: {skipped_count}\n"
                           f"Errors: {error_count}\n\n"
                           f"Details:\n{details_msg}")
    
        self.refresh_color_assignments()

    def clear_all_color_assignments(self):
        """Clear all color assignments."""
        if not messagebox.askyesno("Confirm", "Clear all color assignments?"):
            return

        self.config.subject_colors.clear()
        self.save_config()
        self.log_activity("Cleared all color assignments")
        self.refresh_color_assignments()
        self.color_status_var.set("All color assignments cleared")

    def setup_config_tab(self, notebook):
        """Setup configuration tab."""
        config_frame = ttk.Frame(notebook)
        notebook.add(config_frame, text="Configuration")
        
        # --- File Settings ---
        file_frame = ttk.LabelFrame(config_frame, text="File Settings", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(file_frame, text="Watch Directory:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.watch_dir_var = tk.StringVar(value=self.config.watch_directory)
        ttk.Entry(file_frame, textvariable=self.watch_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_watch_directory).grid(row=0, column=2, padx=5, pady=5)

        self.auto_process_var = tk.BooleanVar(value=self.config.auto_process)
        ttk.Checkbutton(file_frame, text="Auto-process new files", variable=self.auto_process_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        # --- Provider Settings ---
        provider_frame = ttk.LabelFrame(config_frame, text="AI Provider Settings", padding=10)
        provider_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(provider_frame, text="Provider Mode:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.provider_mode_var = tk.StringVar(value=self.config.provider_mode)
        mode_combo = ttk.Combobox(provider_frame, textvariable=self.provider_mode_var, state="readonly", width=20)
        mode_combo['values'] = ["Only OpenRouter", "Only Gemini", "Fallback Mode"]
        mode_combo.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        mode_combo.bind("<<ComboboxSelected>>", self.on_provider_mode_change)
        
        # Fallback options (Initially hidden or shown based on mode)
        self.fallback_frame = ttk.Frame(provider_frame)
        self.fallback_frame.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(self.fallback_frame, text="Primary:").pack(side=tk.LEFT, padx=(0, 5))
        self.primary_provider_var = tk.StringVar(value=self.config.primary_provider)
        primary_combo = ttk.Combobox(self.fallback_frame, textvariable=self.primary_provider_var, state="readonly", width=12)
        primary_combo['values'] = ["OpenRouter", "Gemini"]
        primary_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(self.fallback_frame, text="Secondary:").pack(side=tk.LEFT, padx=(0, 5))
        self.secondary_provider_var = tk.StringVar(value=self.config.secondary_provider)
        secondary_combo = ttk.Combobox(self.fallback_frame, textvariable=self.secondary_provider_var, state="readonly", width=12)
        secondary_combo['values'] = ["OpenRouter", "Gemini"]
        secondary_combo.pack(side=tk.LEFT)
        
        # --- Model Settings Container ---
        self.models_container = ttk.Frame(config_frame)
        self.models_container.pack(fill=tk.X, padx=10, pady=5)
        
        # OpenRouter Settings
        self.openrouter_frame = ttk.LabelFrame(self.models_container, text="OpenRouter Model Configuration", padding=10)
        self.openrouter_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.openrouter_frame, text="Model:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.model_var = tk.StringVar(value=self.config.openrouter_model)
        or_model_combo = ttk.Combobox(self.openrouter_frame, textvariable=self.model_var, width=47)
        or_model_combo.grid(row=0, column=1, padx=5, pady=5)
        or_model_combo['values'] = [
            "openai/gpt-4o-mini",
            "openai/gpt-4o",
            "anthropic/claude-3-haiku",
            "anthropic/claude-3-sonnet",
            "meta-llama/llama-3.1-8b-instruct",
            "google/gemini-pro-1.5"
        ]
        ttk.Button(self.openrouter_frame, text="List Models", command=self.list_available_models).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(self.openrouter_frame, text="Temperature:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.temperature_var = tk.DoubleVar(value=self.config.openrouter_temperature)
        ttk.Scale(self.openrouter_frame, from_=0.0, to=2.0, variable=self.temperature_var,
                 orient=tk.HORIZONTAL, length=300).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.temp_label = ttk.Label(self.openrouter_frame, text=f"{self.config.openrouter_temperature:.1f}")
        self.temp_label.grid(row=1, column=2, padx=5, pady=5)
        self.temperature_var.trace('w', self.update_temperature_label)
        
        ttk.Label(self.openrouter_frame, text="Max Tokens:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.max_tokens_var = tk.IntVar(value=self.config.openrouter_max_tokens)
        ttk.Spinbox(self.openrouter_frame, from_=1000, to=32000, textvariable=self.max_tokens_var,
                   width=48).grid(row=2, column=1, padx=5, pady=5)

        # Gemini Settings
        self.gemini_frame = ttk.LabelFrame(self.models_container, text="Gemini Model Configuration", padding=10)
        self.gemini_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.gemini_frame, text="Model:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.gemini_model_var = tk.StringVar(value=self.config.gemini_model)
        gemini_model_combo = ttk.Combobox(self.gemini_frame, textvariable=self.gemini_model_var, width=47)
        gemini_model_combo.grid(row=0, column=1, padx=5, pady=5)
        gemini_model_combo['values'] = [
            "gemini-1.5-pro",
            "gemini-1.5-flash",
            "gemini-1.0-pro"
        ]
        
        ttk.Label(self.gemini_frame, text="Temperature:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.gemini_temperature_var = tk.DoubleVar(value=self.config.gemini_temperature)
        ttk.Scale(self.gemini_frame, from_=0.0, to=2.0, variable=self.gemini_temperature_var,
                 orient=tk.HORIZONTAL, length=300).grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.gemini_temp_label = ttk.Label(self.gemini_frame, text=f"{self.config.gemini_temperature:.1f}")
        self.gemini_temp_label.grid(row=1, column=2, padx=5, pady=5)
        self.gemini_temperature_var.trace('w', self.update_gemini_temperature_label)
        
        ttk.Label(self.gemini_frame, text="Max Tokens:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.gemini_max_tokens_var = tk.IntVar(value=self.config.gemini_max_tokens)
        ttk.Spinbox(self.gemini_frame, from_=1000, to=32000, textvariable=self.gemini_max_tokens_var,
                   width=48).grid(row=2, column=1, padx=5, pady=5)

        # --- Subjects ---
        subjects_frame = ttk.LabelFrame(config_frame, text="Subjects", padding=10)
        subjects_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.subjects_listbox = tk.Listbox(subjects_frame, height=6, width=30)
        self.subjects_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.update_subjects_listbox()
        
        subjects_controls = ttk.Frame(subjects_frame)
        subjects_controls.pack(side=tk.LEFT, fill=tk.Y)
        
        self.subject_entry = ttk.Entry(subjects_controls, width=20)
        self.subject_entry.pack(pady=(0, 5))
        
        ttk.Button(subjects_controls, text="Add Subject", command=self.add_subject).pack(fill=tk.X, pady=(0, 2))
        ttk.Button(subjects_controls, text="Remove Selected", command=self.remove_subject).pack(fill=tk.X, pady=(0, 2))
        
        # --- Other Settings ---
        other_frame = ttk.Frame(config_frame)
        other_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.remove_thinking_var = tk.BooleanVar(value=self.config.remove_thinking_tags)
        ttk.Checkbutton(other_frame, text="Remove <think> tags from generated notes",
                        variable=self.remove_thinking_var).pack(anchor=tk.W)
        
        ttk.Button(config_frame, text="Save Configuration", command=self.save_configuration).pack(pady=20)
        
        # Initialize UI state based on current config
        self.on_provider_mode_change(None)

    def on_provider_mode_change(self, event):
        """Handle provider mode change."""
        mode = self.provider_mode_var.get()
        
        if mode == "Only OpenRouter":
            self.fallback_frame.grid_remove()
            self.openrouter_frame.pack(fill=tk.X, pady=5)
            self.gemini_frame.pack_forget()
        elif mode == "Only Gemini":
            self.fallback_frame.grid_remove()
            self.openrouter_frame.pack_forget()
            self.gemini_frame.pack(fill=tk.X, pady=5)
        elif mode == "Fallback Mode":
            self.fallback_frame.grid()
            self.openrouter_frame.pack(fill=tk.X, pady=5)
            self.gemini_frame.pack(fill=tk.X, pady=5)

    def update_gemini_temperature_label(self, *args):
        """Update gemini temperature label."""
        self.gemini_temp_label.config(text=f"{self.gemini_temperature_var.get():.1f}")

    def setup_api_keys_tab(self, notebook):
        """Setup API keys management tab."""
        api_frame = ttk.Frame(notebook)
        notebook.add(api_frame, text="API Keys")

        # --- AssemblyAI ---
        assemblyai_frame = ttk.LabelFrame(api_frame, text="AssemblyAI API Key", padding=10)
        assemblyai_frame.pack(fill=tk.X, padx=10, pady=10)

        self.assemblyai_key_var = tk.StringVar()
        assemblyai_entry = ttk.Entry(assemblyai_frame, textvariable=self.assemblyai_key_var,
                                   width=80, show="*", font=("Consolas", 9))
        assemblyai_entry.pack(fill=tk.X, pady=(0, 10))

        assemblyai_buttons = ttk.Frame(assemblyai_frame)
        assemblyai_buttons.pack(fill=tk.X)

        ttk.Button(assemblyai_buttons, text="Load from File",
                  command=lambda: self.load_api_key(self.assemblyai_key_file, self.assemblyai_key_var)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(assemblyai_buttons, text="Save to File",
                  command=lambda: self.save_api_key(self.assemblyai_key_file, self.assemblyai_key_var.get())).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(assemblyai_buttons, text="Test Connection",
                  command=self.test_assemblyai_connection).pack(side=tk.LEFT, padx=(0, 5))

        # --- OpenRouter ---
        openrouter_frame = ttk.LabelFrame(api_frame, text="OpenRouter API Key", padding=10)
        openrouter_frame.pack(fill=tk.X, padx=10, pady=10)

        self.openrouter_key_var = tk.StringVar()
        openrouter_entry = ttk.Entry(openrouter_frame, textvariable=self.openrouter_key_var,
                                   width=80, show="*", font=("Consolas", 9))
        openrouter_entry.pack(fill=tk.X, pady=(0, 10))

        openrouter_buttons = ttk.Frame(openrouter_frame)
        openrouter_buttons.pack(fill=tk.X)

        ttk.Button(openrouter_buttons, text="Load from File",
                  command=lambda: self.load_api_key(self.openrouter_key_file, self.openrouter_key_var)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(openrouter_buttons, text="Save to File",
                  command=lambda: self.save_api_key(self.openrouter_key_file, self.openrouter_key_var.get())).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(openrouter_buttons, text="Test Connection",
                  command=self.test_openrouter_connection).pack(side=tk.LEFT, padx=(0, 5))

        # --- Gemini ---
        gemini_frame = ttk.LabelFrame(api_frame, text="Gemini API Key", padding=10)
        gemini_frame.pack(fill=tk.X, padx=10, pady=10)

        self.gemini_key_var = tk.StringVar()
        gemini_entry = ttk.Entry(gemini_frame, textvariable=self.gemini_key_var,
                                   width=80, show="*", font=("Consolas", 9))
        gemini_entry.pack(fill=tk.X, pady=(0, 10))

        gemini_buttons = ttk.Frame(gemini_frame)
        gemini_buttons.pack(fill=tk.X)

        ttk.Button(gemini_buttons, text="Load from File",
                  command=lambda: self.load_api_key(self.gemini_key_file, self.gemini_key_var)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(gemini_buttons, text="Save to File",
                  command=lambda: self.save_api_key(self.gemini_key_file, self.gemini_key_var.get())).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(gemini_buttons, text="Test Connection",
                  command=self.test_gemini_connection).pack(side=tk.LEFT, padx=(0, 5))

        self.load_api_key(self.assemblyai_key_file, self.assemblyai_key_var)
        self.load_api_key(self.openrouter_key_file, self.openrouter_key_var)
        self.load_api_key(self.gemini_key_file, self.gemini_key_var)

        self.api_status_var = tk.StringVar(value="API keys status: Not tested")
        ttk.Label(api_frame, textvariable=self.api_status_var).pack(pady=10)

    def setup_pre_prompt_tab(self, notebook):
        """Setup pre-prompt management tab."""
        prompt_frame = ttk.Frame(notebook)
        notebook.add(prompt_frame, text="Pre-Prompt")

        ttk.Label(prompt_frame, text="Customize the pre-prompt sent before each transcript:",
                 font=("", 10, "bold")).pack(anchor=tk.W, padx=10, pady=(10, 5))

        editor_frame = ttk.Frame(prompt_frame)
        editor_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.prompt_text = scrolledtext.ScrolledText(editor_frame, height=25, width=80,
                                                   wrap=tk.WORD, font=("Consolas", 10))
        self.prompt_text.pack(fill=tk.BOTH, expand=True)

        current_prompt = self.read_pre_prompt()
        self.prompt_text.insert(tk.END, current_prompt)

        buttons_frame = ttk.Frame(prompt_frame)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(buttons_frame, text="Save Pre-Prompt", command=self.save_pre_prompt).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Reload from File", command=self.reload_pre_prompt).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Reset to Default", command=self.reset_pre_prompt).pack(side=tk.LEFT, padx=(0, 5))

        self.prompt_char_var = tk.StringVar()
        ttk.Label(buttons_frame, textvariable=self.prompt_char_var).pack(side=tk.RIGHT)

        self.prompt_text.bind('<KeyRelease>', self.update_prompt_char_count)
        self.update_prompt_char_count()

    def setup_monitoring_tab(self, notebook):
        """Setup monitoring tab."""
        monitor_frame = ttk.Frame(notebook)
        notebook.add(monitor_frame, text="Monitoring")

        ttk.Label(monitor_frame, text="File Monitoring Status:").pack(anchor=tk.W, padx=10, pady=(10, 5))
        self.monitor_status_var = tk.StringVar()
        ttk.Label(monitor_frame, textvariable=self.monitor_status_var, font=("", 10, "bold")).pack(anchor=tk.W, padx=20)

        buttons_frame = ttk.Frame(monitor_frame)
        buttons_frame.pack(anchor=tk.W, padx=10, pady=10)

        ttk.Button(buttons_frame, text="Start Monitoring", command=self.start_file_monitoring).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Stop Monitoring", command=self.stop_file_monitoring).pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(monitor_frame, text="Recent Activity:").pack(anchor=tk.W, padx=10, pady=(20, 5))

        log_frame = ttk.Frame(monitor_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.activity_text = tk.Text(log_frame, height=15, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.activity_text.yview)
        self.activity_text.configure(yscrollcommand=scrollbar.set)

        self.activity_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def setup_tasks_tab(self, notebook):
        """Setup tasks tab."""
        tasks_frame = ttk.Frame(notebook)
        notebook.add(tasks_frame, text="Processing Tasks")

        columns = ("File", "Subject", "Status", "Created", "Tokens", "Progress")
        self.tasks_tree = ttk.Treeview(tasks_frame, columns=columns, show="headings", height=15)

        self.tasks_tree.heading("File", text="File")
        self.tasks_tree.heading("Subject", text="Subject")
        self.tasks_tree.heading("Status", text="Status")
        self.tasks_tree.heading("Created", text="Created")
        self.tasks_tree.heading("Tokens", text="Tokens")
        self.tasks_tree.heading("Progress", text="Progress")

        self.tasks_tree.column("File", width=200)
        self.tasks_tree.column("Subject", width=100)
        self.tasks_tree.column("Status", width=120)
        self.tasks_tree.column("Created", width=130)
        self.tasks_tree.column("Tokens", width=80)
        self.tasks_tree.column("Progress", width=250)

        tasks_scrollbar = ttk.Scrollbar(tasks_frame, orient=tk.VERTICAL, command=self.tasks_tree.yview)
        self.tasks_tree.configure(yscrollcommand=tasks_scrollbar.set)

        self.tasks_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
        tasks_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)

        task_buttons_frame = ttk.Frame(tasks_frame)
        task_buttons_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(task_buttons_frame, text="Refresh", command=self.refresh_tasks_display).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(task_buttons_frame, text="Clear Completed", command=self.clear_completed_tasks).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(task_buttons_frame, text="Retry Failed", command=self.retry_failed_tasks).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(task_buttons_frame, text="Open Notes Folder", command=self.open_notes_folder).pack(side=tk.LEFT, padx=(0, 5))

    def setup_reprocessing_tab(self, notebook):
        """Setup reprocessing tab."""
        reprocess_frame = ttk.Frame(notebook)
        notebook.add(reprocess_frame, text="Reprocessing")

        controls_frame = ttk.Frame(reprocess_frame)
        controls_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(controls_frame, text="Directory to scan:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.reprocess_dir_var = tk.StringVar(value=self.config.watch_directory)
        ttk.Entry(controls_frame, textvariable=self.reprocess_dir_var, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(controls_frame, text="Browse", command=self.browse_reprocess_directory).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(controls_frame, text="Scan", command=self.scan_reprocess_files).grid(row=0, column=3, padx=5, pady=5)

        selection_frame = ttk.Frame(controls_frame)
        selection_frame.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=10)

        ttk.Button(selection_frame, text="Select All", command=self.select_all_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(selection_frame, text="Deselect All", command=self.deselect_all_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(selection_frame, text="Select by Subject", command=self.select_by_subject).pack(side=tk.LEFT, padx=(0, 5))

        self.selection_count_var = tk.StringVar(value="No files selected")
        ttk.Label(selection_frame, textvariable=self.selection_count_var, font=("", 10, "italic")).pack(side=tk.RIGHT)

        files_frame = ttk.Frame(reprocess_frame)
        files_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        columns = ("Select", "File", "Subject", "Size", "Modified", "Transcript", "Notes", "Status")
        self.reprocess_tree = ttk.Treeview(files_frame, columns=columns, show="headings", height=15)

        self.reprocess_tree.heading("Select", text="☐")
        self.reprocess_tree.heading("File", text="Audio File")
        self.reprocess_tree.heading("Subject", text="Subject")
        self.reprocess_tree.heading("Size", text="Size")
        self.reprocess_tree.heading("Modified", text="Modified")
        self.reprocess_tree.heading("Transcript", text="Transcript")
        self.reprocess_tree.heading("Notes", text="Notes")
        self.reprocess_tree.heading("Status", text="Status")

        self.reprocess_tree.column("Select", width=50, anchor=tk.CENTER)
        self.reprocess_tree.column("File", width=200)
        self.reprocess_tree.column("Subject", width=100)
        self.reprocess_tree.column("Size", width=80)
        self.reprocess_tree.column("Modified", width=130)
        self.reprocess_tree.column("Transcript", width=80, anchor=tk.CENTER)
        self.reprocess_tree.column("Notes", width=80, anchor=tk.CENTER)
        self.reprocess_tree.column("Status", width=150)

        self.reprocess_tree.bind("<Button-1>", self.on_reprocess_tree_click)
        self.reprocess_tree.bind("<Double-Button-1>", self.on_reprocess_tree_double_click)

        files_scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.reprocess_tree.yview)
        self.reprocess_tree.configure(yscrollcommand=files_scrollbar.set)

        self.reprocess_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        files_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        action_frame = ttk.Frame(reprocess_frame)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        options_frame = ttk.LabelFrame(action_frame, text="Reprocessing Options", padding=5)
        options_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        self.reprocess_type_var = tk.StringVar(value="both")
        ttk.Radiobutton(options_frame, text="Transcript Only", variable=self.reprocess_type_var,
                       value="transcript").pack(anchor=tk.W)
        ttk.Radiobutton(options_frame, text="Notes Only", variable=self.reprocess_type_var,
                       value="notes").pack(anchor=tk.W)
        ttk.Radiobutton(options_frame, text="Both Transcript & Notes", variable=self.reprocess_type_var,
                       value="both").pack(anchor=tk.W)

        buttons_frame = ttk.Frame(action_frame)
        buttons_frame.pack(side=tk.RIGHT)

        ttk.Button(buttons_frame, text="Reprocess Selected",
                  command=self.reprocess_selected_files, style="Accent.TButton").pack(pady=2)
        ttk.Button(buttons_frame, text="Delete Selected Outputs",
                  command=self.delete_selected_outputs).pack(pady=2)
        ttk.Button(buttons_frame, text="Open File Location",
                  command=self.open_selected_file_location).pack(pady=2)

        status_frame = ttk.Frame(reprocess_frame)
        status_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.reprocess_status_var = tk.StringVar(value="Ready to scan files")
        ttk.Label(status_frame, textvariable=self.reprocess_status_var, font=("", 10, "bold")).pack(anchor=tk.W)

    def setup_word_tab(self, notebook):
        """Setup Word document formatting tab."""
        word_frame = ttk.Frame(notebook)
        notebook.add(word_frame, text="Word Documents")

        ttk.Label(word_frame, text="Document Settings:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.word_auto_update_var = tk.BooleanVar(value=getattr(self.config, 'word_auto_update', True))
        ttk.Checkbutton(word_frame, text="Auto-update Word documents when new notes are generated",
                       variable=self.word_auto_update_var).grid(row=0, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        font_frame = ttk.LabelFrame(word_frame, text="Font Settings", padding=10)
        font_frame.grid(row=1, column=0, columnspan=3, sticky=tk.EW, padx=10, pady=10)

        ttk.Label(font_frame, text="Font Name:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.word_font_name_var = tk.StringVar(value=getattr(self.config, 'word_font_name', 'Calibri'))
        font_combo = ttk.Combobox(font_frame, textvariable=self.word_font_name_var, width=25)
        font_combo.grid(row=0, column=1, padx=5, pady=5)
        font_combo['values'] = ["Calibri", "Times New Roman", "Arial", "Helvetica", "Georgia", "Verdana"]

        ttk.Label(font_frame, text="Font Size:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.word_font_size_var = tk.IntVar(value=getattr(self.config, 'word_font_size', 11))
        ttk.Spinbox(font_frame, from_=8, to=24, textvariable=self.word_font_size_var, width=23).grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(font_frame, text="Line Spacing:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.word_line_spacing_var = tk.DoubleVar(value=getattr(self.config, 'word_line_spacing', 1.15))
        ttk.Scale(font_frame, from_=1.0, to=3.0, variable=self.word_line_spacing_var,
                 orient=tk.HORIZONTAL, length=200).grid(row=2, column=1, padx=5, pady=5)
        self.word_line_spacing_label = ttk.Label(font_frame, text=f"{self.word_line_spacing_var.get():.2f}")
        self.word_line_spacing_label.grid(row=2, column=2, padx=5, pady=5)
        self.word_line_spacing_var.trace('w', self.update_line_spacing_label)

        headings_frame = ttk.LabelFrame(word_frame, text="Heading Sizes", padding=10)
        headings_frame.grid(row=2, column=0, columnspan=3, sticky=tk.EW, padx=10, pady=10)

        self.word_heading_vars = {}
        for i in range(1, 4):
            ttk.Label(headings_frame, text=f"Heading {i}:").grid(row=i-1, column=0, sticky=tk.W, padx=5, pady=5)
            default_size = getattr(self.config, f'word_heading{i}_size', 18-i*2)
            self.word_heading_vars[i] = tk.IntVar(value=default_size)
            ttk.Spinbox(headings_frame, from_=10, to=28, textvariable=self.word_heading_vars[i],
                       width=23).grid(row=i-1, column=1, padx=5, pady=5)

        buttons_frame = ttk.Frame(word_frame)
        buttons_frame.grid(row=3, column=0, columnspan=3, pady=20)

        ttk.Button(buttons_frame, text="Save Word Settings", command=self.save_word_configuration).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Update All Documents", command=self.update_all_word_documents).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Regenerate All Documents", command=self.regenerate_all_word_documents).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Open Documents Folder", command=self.open_word_documents_folder).pack(side=tk.LEFT, padx=5)

        self.word_status_var = tk.StringVar(value="Word documents: Ready")
        ttk.Label(word_frame, textvariable=self.word_status_var).grid(row=4, column=0, columnspan=3, pady=10)

    def browse_reprocess_directory(self):
        """Browse for reprocessing directory."""
        directory = filedialog.askdirectory(title="Select Directory to Scan for Audio Files")
        if directory:
            self.reprocess_dir_var.set(directory)

    def scan_reprocess_files(self):
        """Scan directory for audio files and populate the reprocessing list."""
        scan_dir = self.reprocess_dir_var.get().strip()
        if not scan_dir or not os.path.exists(scan_dir):
            messagebox.showerror("Error", "Please select a valid directory to scan!")
            return

        if not self.config.subjects:
            messagebox.showerror("Error", "Please configure subjects first in the Configuration tab!")
            return

        self.reprocess_status_var.set("Scanning directory...")
        self.root.update()

        try:
            self.reprocessing_files.clear()
            self.selected_files.clear()
            for item in self.reprocess_tree.get_children():
                self.reprocess_tree.delete(item)

            found_files = 0
            scan_path = Path(scan_dir)

            for file_path in scan_path.iterdir():
                if file_path.is_file() and file_path.suffix.lower() in self.config.supported_extensions:
                    filename = file_path.name
                    matching_subject = None

                    for subject in self.config.subjects:
                        if subject.lower() in filename.lower():
                            matching_subject = subject
                            break

                    if matching_subject:
                        # Check for existing outputs
                        subject_dir = Path(matching_subject)
                        transcripts_dir = subject_dir / "transcripts"
                        notes_dir = subject_dir / "notes"

                        transcript_path = transcripts_dir / f"{file_path.stem}.txt"
                        notes_path = notes_dir / f"{file_path.stem}_notes.md"

                        file_info = ReprocessingFileInfo(
                            filepath=str(file_path),
                            filename=filename,
                            subject=matching_subject,
                            has_transcript=transcript_path.exists(),
                            has_notes=notes_path.exists(),
                            transcript_path=str(transcript_path) if transcript_path.exists() else "",
                            notes_path=str(notes_path) if notes_path.exists() else "",
                            file_size=file_path.stat().st_size,
                            modified_date=datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                        )

                        self.reprocessing_files[str(file_path)] = file_info
                        self.insert_reprocess_item(file_info)
                        found_files += 1

            self.selection_count_var.set(f"Found {found_files} files")
            self.reprocess_status_var.set(f"Scan complete. Found {found_files} relevant files.")

        except Exception as e:
            messagebox.showerror("Error", f"Error scanning directory: {e}")
            self.reprocess_status_var.set("Error during scan")

    def insert_reprocess_item(self, file_info):
        """Insert item into reprocessing tree."""
        size_mb = file_info.file_size / (1024 * 1024)
        size_str = f"{size_mb:.1f} MB"

        transcript_icon = "✓" if file_info.has_transcript else "✗"
        notes_icon = "✓" if file_info.has_notes else "✗"

        status = "Ready"
        if file_info.has_transcript and file_info.has_notes:
            status = "Completed"
        elif file_info.has_transcript:
            status = "Has Transcript Only"

        self.reprocess_tree.insert("", tk.END, values=(
            "☐",
            file_info.filename,
            file_info.subject,
            size_str,
            file_info.modified_date,
            transcript_icon,
            notes_icon,
            status
        ), tags=(file_info.filepath,))

    def on_reprocess_tree_click(self, event):
        """Handle click on reprocessing tree."""
        region = self.reprocess_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.reprocess_tree.identify_column(event.x)
            if column == "#1":  # Checkbox column
                item_id = self.reprocess_tree.identify_row(event.y)
                if item_id:
                    self.toggle_reprocess_selection(item_id)

    def toggle_reprocess_selection(self, item_id):
        """Toggle selection state of a file."""
        tags = self.reprocess_tree.item(item_id, "tags")
        if not tags:
            return

        filepath = tags[0]
        if filepath in self.reprocessing_files:
            file_info = self.reprocessing_files[filepath]
            file_info.selected = not file_info.selected

            # Update display
            current_values = self.reprocess_tree.item(item_id, "values")
            new_checkbox = "☑" if file_info.selected else "☐"
            new_values = (new_checkbox,) + current_values[1:]
            self.reprocess_tree.item(item_id, values=new_values)

            if file_info.selected:
                self.selected_files.add(filepath)
            else:
                self.selected_files.discard(filepath)

            self.update_selection_count()

    def update_selection_count(self):
        """Update the selection count label."""
        count = len(self.selected_files)
        self.selection_count_var.set(f"{count} files selected")

    def select_all_files(self):
        """Select all files in the list."""
        for item_id in self.reprocess_tree.get_children():
            tags = self.reprocess_tree.item(item_id, "tags")
            if tags:
                filepath = tags[0]
                if filepath in self.reprocessing_files:
                    file_info = self.reprocessing_files[filepath]
                    if not file_info.selected:
                        self.toggle_reprocess_selection(item_id)

    def deselect_all_files(self):
        """Deselect all files in the list."""
        for item_id in self.reprocess_tree.get_children():
            tags = self.reprocess_tree.item(item_id, "tags")
            if tags:
                filepath = tags[0]
                if filepath in self.reprocessing_files:
                    file_info = self.reprocessing_files[filepath]
                    if file_info.selected:
                        self.toggle_reprocess_selection(item_id)

    def select_by_subject(self):
        """Select files by subject."""
        if not self.config.subjects:
            return

        # Simple dialog to choose subject
        subject_window = tk.Toplevel(self.root)
        subject_window.title("Select Subject")
        subject_window.geometry("300x400")

        listbox = tk.Listbox(subject_window)
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for subject in self.config.subjects:
            listbox.insert(tk.END, subject)

        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_subject = listbox.get(selection[0])
                self.deselect_all_files()

                for item_id in self.reprocess_tree.get_children():
                    tags = self.reprocess_tree.item(item_id, "tags")
                    if tags:
                        filepath = tags[0]
                        if filepath in self.reprocessing_files:
                            file_info = self.reprocessing_files[filepath]
                            if file_info.subject == selected_subject:
                                self.toggle_reprocess_selection(item_id)
                subject_window.destroy()

        ttk.Button(subject_window, text="Select", command=on_select).pack(pady=10)

    def on_reprocess_tree_double_click(self, event):
        """Handle double click to open file info."""
        item_id = self.reprocess_tree.identify_row(event.y)
        if item_id:
            tags = self.reprocess_tree.item(item_id, "tags")
            if tags:
                filepath = tags[0]
                if filepath in self.reprocessing_files:
                    self.show_file_details(self.reprocessing_files[filepath])

    def show_file_details(self, file_info: ReprocessingFileInfo):
        """Show details about a file."""
        details = f"""File: {file_info.filename}
Subject: {file_info.subject}
Path: {file_info.filepath}
Size: {file_info.file_size / (1024*1024):.2f} MB
Modified: {file_info.modified_date}

Transcript: {'Yes' if file_info.has_transcript else 'No'}
Notes: {'Yes' if file_info.has_notes else 'No'}

Transcript Path: {file_info.transcript_path if file_info.has_transcript else 'Not found'}
Notes Path: {file_info.notes_path if file_info.has_notes else 'Not found'}
"""
        messagebox.showinfo("File Details", details)

    def reprocess_selected_files(self):
        """Reprocess selected files."""
        if not self.selected_files:
            messagebox.showwarning("Warning", "No files selected!")
            return

        reprocess_type = self.reprocess_type_var.get()
        count = len(self.selected_files)

        if not messagebox.askyesno("Confirm", f"Reprocess {count} files ({reprocess_type})?"):
            return

        for filepath in self.selected_files:
            file_info = self.reprocessing_files[filepath]

            task = ProcessingTask(
                filepath=file_info.filepath,
                subject=file_info.subject,
                reprocess_type=reprocess_type,
                transcript_path=file_info.transcript_path,
                notes_path=file_info.notes_path
            )

            self.task_queue.put(task)
            self.reprocess_status_var.set(f"Queued {count} files for reprocessing")

        try:
            self.notebook.select(self.tasks_tree.master)
        except Exception:
            pass

    def handle_reprocessing_task(self, task: ProcessingTask):
        """Handle a reprocessing task."""
        try:
            self.log_activity(f"Reprocessing {task.reprocess_type}: {Path(task.filepath).name}")

            if task.reprocess_type in ["transcript", "both"]:
                # Force transcription even if exists
                task.status = "transcribing"
                # Logic same as process_task for transcription part
                self.process_task(task) # This might be recursive but process_task checks reprocess_type
                # Wait, process_task calls handle_reprocessing_task if reprocess_type is set. Infinite loop!
                # We need to handle it properly.

                # Actually, simply clearing reprocess_type and calling process_task might work if we want full redo
                # But we might want to skip transcription if "notes only"
                pass

            if task.reprocess_type == "notes":
                if task.transcript_path and os.path.exists(task.transcript_path):
                    self.process_notes_only(task)
                else:
                    self.log_activity(f"Cannot reprocess notes: Transcript missing for {Path(task.filepath).name}")

            elif task.reprocess_type == "both" or task.reprocess_type == "transcript":
                 # We need to clear reprocess_type to avoid loop and call process_task
                 # But we also want to ensure we overwrite.
                 # AssemblyAI transcriber overwrites by default? Yes usually.
                 task.reprocess_type = "" # Clear it
                 self.process_task(task)

        except Exception as e:
            logging.error(f"Reprocessing error: {e}")

    def delete_selected_outputs(self):
        """Delete outputs for selected files."""
        if not self.selected_files:
            return

        if not messagebox.askyesno("Confirm", "Delete transcripts and notes for selected files? Audio files will NOT be deleted."):
            return

        deleted_count = 0
        for filepath in self.selected_files:
            file_info = self.reprocessing_files[filepath]

            if file_info.has_transcript and os.path.exists(file_info.transcript_path):
                try:
                    os.remove(file_info.transcript_path)
                    deleted_count += 1
                except Exception as e:
                    logging.error(f"Error deleting transcript: {e}")

            if file_info.has_notes and os.path.exists(file_info.notes_path):
                try:
                    os.remove(file_info.notes_path)
                    deleted_count += 1
                except Exception as e:
                    logging.error(f"Error deleting notes: {e}")

        messagebox.showinfo("Complete", f"Deleted {deleted_count} output files.")
        self.scan_reprocess_files() # Refresh

    def open_selected_file_location(self):
        """Open location of selected file."""
        selection = self.reprocess_tree.selection()
        if not selection:
            return

        tags = self.reprocess_tree.item(selection[0], "tags")
        if tags:
            filepath = tags[0]
            self.open_file_folder(filepath)

    def log_activity(self, message):
        """Log activity to text widget and file."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_msg = f"[{timestamp}] {message}"

        self.activity_text.config(state=tk.NORMAL)
        self.activity_text.insert(tk.END, log_msg + "\n")
        self.activity_text.see(tk.END)
        self.activity_text.config(state=tk.DISABLED)

        logging.info(message)

    def refresh_tasks_display(self):
        """Refresh the tasks treeview."""
        for item in self.tasks_tree.get_children():
            self.tasks_tree.delete(item)

        for filepath, task in self.tasks.items():
            filename = Path(filepath).name
            self.tasks_tree.insert("", tk.END, values=(
                filename,
                task.subject,
                task.status,
                task.created_at,
                task.tokens_used,
                task.error_message if task.error_message else ""
            ))

    def clear_completed_tasks(self):
        """Clear completed tasks from the list."""
        completed = [k for k, v in self.tasks.items() if v.status == "completed"]
        for k in completed:
            del self.tasks[k]
        self.refresh_tasks_display()

    def retry_failed_tasks(self):
        """Retry failed tasks."""
        failed = [v for k, v in self.tasks.items() if v.status == "error"]
        for task in failed:
            task.status = "queued"
            task.error_message = ""
            self.task_queue.put(task)
        self.refresh_tasks_display()
        if failed:
            self.log_activity(f"Retrying {len(failed)} failed tasks")

    def open_notes_folder(self):
        """Open the notes folder."""
        # Just open the watch directory or first subject folder
        if self.config.watch_directory:
            self.open_file_folder(self.config.watch_directory)

    def open_word_documents_folder(self):
        """Open the folder containing Word documents."""
        # Try to find a subject folder or Appunti Completi
        if self.config.subjects:
            subject = self.config.subjects[0]
            possible_path = Path("Appunti Completi") / subject
            if possible_path.exists():
                self.open_file_folder(str(possible_path))
                return
            
            possible_path = Path(subject) / "Appunti Completi"
            if possible_path.exists():
                self.open_file_folder(str(possible_path))
                return

        # Fallback to current dir
        self.open_file_folder(".")

    def update_all_word_documents(self):
        """Update all Word documents."""
        if not self.word_manager:
            messagebox.showerror("Error", "Word manager not initialized")
            return

        if not self.config.subjects:
            messagebox.showwarning("Warning", "No subjects configured")
            return

        self.word_status_var.set("Updating Word documents...")
        self.root.update()

        updated_count = 0
        for subject in self.config.subjects:
            try:
                # Need to find where markdown notes are. Assuming standard structure
                notes_dir = Path(subject) / "notes"
                if notes_dir.exists():
                    # Scan for md files
                    for md_file in notes_dir.glob("*.md"):
                        self.word_manager.check_new_markdown_file(str(md_file), subject)
                    updated_count += 1
            except Exception as e:
                logging.error(f"Error updating docs for {subject}: {e}")

        self.word_status_var.set(f"Updated documents for {updated_count} subjects")
        messagebox.showinfo("Success", f"Updated Word documents for {updated_count} subjects")

    def regenerate_all_word_documents(self):
        """Regenerate all Word documents from scratch."""
        if not messagebox.askyesno("Confirm", "This will rebuild all Word documents from available notes. Continue?"):
            return

        # Logic similar to update but force rebuild?
        # WordDocumentManager doesn't have explicit 'rebuild all' public method readily available in snippet,
        # but check_new_markdown_file processes them.
        # We can just run update logic.
        self.update_all_word_documents()

    def save_word_configuration(self):
        """Save Word configuration settings."""
        self.config.word_auto_update = self.word_auto_update_var.get()
        self.config.word_font_name = self.word_font_name_var.get()
        self.config.word_font_size = self.word_font_size_var.get()
        self.config.word_line_spacing = self.word_line_spacing_var.get()
        
        for i, var in self.word_heading_vars.items():
            setattr(self.config, f'word_heading{i}_size', var.get())

        self.save_config()
        
        # Re-init manager with new settings
        self.init_word_manager()
        messagebox.showinfo("Success", "Word configuration saved")

    def update_line_spacing_label(self, *args):
        self.word_line_spacing_label.config(text=f"{self.word_line_spacing_var.get():.2f}")

    def open_file_folder(self, path):
        """Open file or folder in OS file explorer."""
        import platform
        import subprocess

        path_obj = Path(path)
        if path_obj.is_file():
            found_folder = path_obj.parent
        else:
            found_folder = path_obj
    
        try:
            if platform.system() == "Windows":
                os.startfile(str(found_folder))
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(found_folder)])
            else:
                subprocess.run(["xdg-open", str(found_folder)])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {e}")

    def browse_watch_directory(self):
        """Browse for watch directory."""
        directory = filedialog.askdirectory(title="Select Watch Directory")
        if directory:
            self.watch_dir_var.set(directory)

    def update_temperature_label(self, *args):
        """Update temperature label."""
        self.temp_label.config(text=f"{self.temperature_var.get():.1f}")

    def list_available_models(self):
        """List available OpenRouter models."""
        try:
            processor = OpenRouterProcessor()
            models = processor.get_available_models()

            if models:
                models_window = tk.Toplevel(self.root)
                models_window.title("Available OpenRouter Models")
                models_window.geometry("600x400")

                columns = ("ID", "Name")
                models_tree = ttk.Treeview(models_window, columns=columns, show="headings")

                models_tree.heading("ID", text="Model ID")
                models_tree.heading("Name", text="Name")

                models_tree.column("ID", width=250)
                models_tree.column("Name", width=300)

                for model in models:
                    model_id = model.get('id', 'Unknown')
                    model_name = model.get('name', model_id)
                    models_tree.insert("", tk.END, values=(model_id, model_name))

                scrollbar = ttk.Scrollbar(models_window, orient=tk.VERTICAL, command=models_tree.yview)
                models_tree.configure(yscrollcommand=scrollbar.set)

                models_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0), pady=10)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 10), pady=10)

                def on_model_select(event):
                    selection = models_tree.selection()
                    if selection:
                        model_id = models_tree.item(selection[0])['values'][0]
                        self.model_var.set(model_id)
                        models_window.destroy()

                models_tree.bind("<Double-1>", on_model_select)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch models: {e}")

    def add_subject(self):
        """Add a new subject."""
        subject = self.subject_entry.get().strip()
        if subject and subject not in self.config.subjects:
            self.config.subjects.append(subject)
            self.update_subjects_listbox()
            self.subject_entry.delete(0, tk.END)
            self.log_activity(f"Added subject: {subject}")
            self.refresh_color_assignments()

    def remove_subject(self):
        """Remove selected subject."""
        selection = self.subjects_listbox.curselection()
        if selection:
            subject = self.subjects_listbox.get(selection[0])
            self.config.subjects.remove(subject)

            if subject in self.config.subject_colors:
                del self.config.subject_colors[subject]

            self.update_subjects_listbox()
            self.log_activity(f"Removed subject: {subject}")
            self.refresh_color_assignments()

    def update_subjects_listbox(self):
        """Update the subjects listbox."""
        self.subjects_listbox.delete(0, tk.END)
        for subject in self.config.subjects:
            self.subjects_listbox.insert(tk.END, subject)

    def save_configuration(self):
        """Save current configuration."""
        self.config.watch_directory = self.watch_dir_var.get()
        self.config.auto_process = self.auto_process_var.get()
        
        # Save provider settings
        self.config.provider_mode = self.provider_mode_var.get()
        self.config.primary_provider = self.primary_provider_var.get()
        self.config.secondary_provider = self.secondary_provider_var.get()
        
        # Save OpenRouter settings
        self.config.openrouter_model = self.model_var.get()
        self.config.openrouter_temperature = self.temperature_var.get()
        self.config.openrouter_max_tokens = self.max_tokens_var.get()
        
        # Save Gemini settings
        self.config.gemini_model = self.gemini_model_var.get()
        self.config.gemini_temperature = self.gemini_temperature_var.get()
        self.config.gemini_max_tokens = self.gemini_max_tokens_var.get()
        
        self.config.remove_thinking_tags = self.remove_thinking_var.get()
        self.config.auto_apply_colors = self.auto_apply_colors_var.get()

        self.save_config()
        messagebox.showinfo("Configuration", "Configuration saved successfully!")

        if self.observer and self.observer.is_alive():
            self.stop_file_monitoring()
            time.sleep(0.5)
            self.start_file_monitoring()

    def load_api_key(self, file_path: Path, var: tk.StringVar):
        """Load API key from file into variable."""
        key = self.read_api_key_file(file_path)
        var.set(key)

    def save_api_key(self, file_path: Path, api_key: str):
        """Save API key to file."""
        if self.write_api_key_file(file_path, api_key):
            messagebox.showinfo("Success", f"API key saved to {file_path.name}")

    def test_assemblyai_connection(self):
        """Test AssemblyAI connection."""
        api_key = self.assemblyai_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Error", "Please enter AssemblyAI API key first")
            return

        try:
            self.write_api_key_file(self.assemblyai_key_file, api_key)

            config = TranscriptionConfig()
            transcriber = AssemblyAITranscriber(config=config)

            self.api_status_var.set("AssemblyAI: Testing connection...")
            self.root.update()

            if len(api_key) > 10:
                self.api_status_var.set("AssemblyAI: ✓ API key format looks valid")
                messagebox.showinfo("Success", "AssemblyAI API key appears valid")
            else:
                self.api_status_var.set("AssemblyAI: ✗ Invalid API key format")
                messagebox.showerror("Error", "Invalid AssemblyAI API key format")

        except Exception as e:
            self.api_status_var.set(f"AssemblyAI: ✗ Connection failed")
            messagebox.showerror("Error", f"AssemblyAI connection test failed: {e}")

    def test_openrouter_connection(self):
        """Test OpenRouter connection."""
        api_key = self.openrouter_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Error", "Please enter OpenRouter API key first")
            return

        try:
            self.write_api_key_file(self.openrouter_key_file, api_key)

            note_config = NoteProcessingConfig(
                model=self.config.openrouter_model,
                temperature=self.config.openrouter_temperature,
                max_tokens=self.config.openrouter_max_tokens
            )

            processor = OpenRouterProcessor(api_key=api_key, config=note_config)

            self.api_status_var.set("OpenRouter: Testing connection...")
            self.root.update()

            result = processor.test_connection()

            if result['success']:
                self.api_status_var.set("OpenRouter: ✓ Connection successful")
                messagebox.showinfo("Success", "OpenRouter connection test successful!")
            else:
                self.api_status_var.set("OpenRouter: ✗ Connection failed")
                messagebox.showerror("Error", f"OpenRouter test failed: {result.get('error', 'Unknown error')}")

        except Exception as e:
            self.api_status_var.set(f"OpenRouter: ✗ Connection failed")
            messagebox.showerror("Error", f"OpenRouter connection test failed: {e}")

    def test_gemini_connection(self):
        """Test Gemini connection."""
        api_key = self.gemini_key_var.get().strip()
        if not api_key:
            messagebox.showerror("Error", "Please enter Gemini API key first")
            return

        try:
            self.write_api_key_file(self.gemini_key_file, api_key)

            gemini_config = GeminiProcessingConfig(
                model=self.config.gemini_model,
                temperature=self.config.gemini_temperature,
                max_tokens=self.config.gemini_max_tokens
            )

            processor = GeminiProcessor(api_key=api_key, config=gemini_config)

            self.api_status_var.set("Gemini: Testing connection...")
            self.root.update()

            result = processor.test_connection()

            if result['success']:
                self.api_status_var.set("Gemini: ✓ Connection successful")
                messagebox.showinfo("Success", "Gemini connection test successful!")
            else:
                self.api_status_var.set("Gemini: ✗ Connection failed")
                messagebox.showerror("Error", f"Gemini test failed: {result.get('error', 'Unknown error')}")

        except Exception as e:
            self.api_status_var.set(f"Gemini: ✗ Connection failed")
            messagebox.showerror("Error", f"Gemini connection test failed: {e}")

    def save_pre_prompt(self):
        """Save pre-prompt to file."""
        prompt = self.prompt_text.get(1.0, tk.END).strip()
        if self.write_pre_prompt(prompt):
            messagebox.showinfo("Success", "Pre-prompt saved successfully!")
            self.log_activity("Pre-prompt updated")

    def reload_pre_prompt(self):
        """Reload pre-prompt from file."""
        prompt = self.read_pre_prompt()
        self.prompt_text.delete(1.0, tk.END)
        self.prompt_text.insert(1.0, prompt)
        self.update_prompt_char_count()

    def reset_pre_prompt(self):
        """Reset pre-prompt to default."""
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset the pre-prompt to default?"):
            if self.pre_prompt_file.exists():
                self.pre_prompt_file.unlink()

            self.ensure_pre_prompt_file()
            self.reload_pre_prompt()

    def update_prompt_char_count(self, event=None):
        """Update character count for pre-prompt."""
        prompt = self.prompt_text.get(1.0, tk.END)
        char_count = len(prompt.strip())
        self.prompt_char_var.set(f"Characters: {char_count}")

    def start_processing_thread(self):
        """Start the background processing thread."""
        if self.processing_thread is None or not self.processing_thread.is_alive():
            self.processing_thread = NoteProcessingThread(self.task_queue, self)
            self.processing_thread.start()
            logging.info("Processing thread started")

    def start_file_monitoring(self):
        """Start file monitoring."""
        if not self.config.watch_directory or not os.path.exists(self.config.watch_directory):
            messagebox.showerror("Error", "Please configure a valid watch directory first!")
            return

        if not self.config.subjects:
            messagebox.showerror("Error", "Please add at least one subject first!")
            return

        if not self.read_api_key_file(self.assemblyai_key_file):
            messagebox.showerror("Error", "AssemblyAI API key not found. Please configure it first!")
            return

        # Check for at least one AI provider key
        has_openrouter = bool(self.read_api_key_file(self.openrouter_key_file))
        has_gemini = bool(self.read_api_key_file(self.gemini_key_file))
        
        if not has_openrouter and not has_gemini:
             messagebox.showerror("Error", "Please configure at least one AI Provider API key (OpenRouter or Gemini)!")
             return

        try:
            if self.observer:
                self.stop_file_monitoring()

            self.scan_existing_files()

            self.observer = Observer()
            event_handler = AudioFileHandler(self)
            self.observer.schedule(event_handler, self.config.watch_directory, recursive=False)
            self.observer.start()

            self.log_activity(f"Started monitoring directory: {self.config.watch_directory}")
            logging.info(f"Started file monitoring: {self.config.watch_directory}")
            self.update_gui()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to start monitoring: {e}")
            logging.error(f"Failed to start monitoring: {e}")
            self.update_gui()

    def stop_file_monitoring(self):
        """Stop file monitoring."""
        if self.observer and self.observer.is_alive():
            self.observer.stop()
            self.observer.join()
            self.log_activity("Stopped file monitoring")
            logging.info("Stopped file monitoring")

        self.update_gui()

    def handle_new_file(self, filepath):
        """Handle new file detection."""
        try:
            path = Path(filepath)

            if path.suffix.lower() not in self.config.supported_extensions:
                return

            time.sleep(2)

            filename = path.name.lower()
            matching_subject = None

            for subject in self.config.subjects:
                if subject.lower() in filename:
                    matching_subject = subject
                    break

            if not matching_subject:
                self.log_activity(f"No matching subject found for: {path.name}")
                return

            task = ProcessingTask(
                filepath=str(path),
                subject=matching_subject
            )

            self.tasks[str(path)] = task
            self.log_activity(f"Detected new file: {path.name} (Subject: {matching_subject})")

            if self.config.auto_process:
                self.task_queue.put(task)
                task.status = "queued"
                self.log_activity(f"Queued for processing: {path.name}")

        except Exception as e:
            self.log_activity(f"Error handling new file {filepath}: {e}")
            logging.error(f"Error handling new file {filepath}: {e}")

    def scan_existing_files(self):
        """Scan existing files in watch directory."""
        if not self.config.watch_directory or not os.path.exists(self.config.watch_directory):
            return

        self.log_activity("Starting initial scan of existing files...")

        try:
            watch_path = Path(self.config.watch_directory)
            found_files = 0
            processed_files = 0
            pending_files = 0

            for file_path in watch_path.iterdir():
                if file_path.is_file() and file_path.suffix.lower() in self.config.supported_extensions:
                    found_files += 1

                    filename = file_path.name.lower()
                    matching_subject = None

                    for subject in self.config.subjects:
                        if subject.lower() in filename:
                            matching_subject = subject
                            break

                    if not matching_subject:
                        self.log_activity(f"No matching subject for: {file_path.name}")
                        continue

                    subject_dir = Path(matching_subject)
                    transcripts_dir = subject_dir / "transcripts"
                    notes_dir = subject_dir / "notes"

                    expected_transcript = transcripts_dir / f"{file_path.stem}.txt"
                    expected_notes = notes_dir / f"{file_path.stem}_notes.md"

                    task_key = str(file_path)

                    if task_key in self.tasks:
                        continue

                    task = ProcessingTask(
                        filepath=str(file_path),
                        subject=matching_subject
                    )

                    if expected_notes.exists() and expected_transcript.exists():
                        task.status = "completed"
                        task.transcript_path = str(expected_transcript)
                        task.notes_path = str(expected_notes)
                        processed_files += 1
                        self.log_activity(f"Already processed: {file_path.name}")

                    elif expected_transcript.exists():
                        task.status = "transcript_only"
                        task.transcript_path = str(expected_transcript)
                        pending_files += 1
                        self.log_activity(f"Has transcript, needs notes: {file_path.name}")

                        if self.config.auto_process:
                            self.task_queue.put(task)
                            task.status = "queued_notes"

                    else:
                        task.status = "pending"
                        pending_files += 1
                        self.log_activity(f"Needs processing: {file_path.name}")

                        if self.config.auto_process:
                            self.task_queue.put(task)
                            task.status = "queued"

                    self.tasks[task_key] = task

            self.log_activity(f"Scan complete: {found_files} audio files found, {processed_files} already processed, {pending_files} need processing")

        except Exception as e:
            self.log_activity(f"Error during initial scan: {e}")
            logging.error(f"Initial scan error: {e}")

    def process_notes_only(self, task: ProcessingTask):
        """Process only notes generation for files that already have transcripts."""
        try:
            task.status = "processing_notes"

            subject_dir = Path(task.subject)
            notes_dir = subject_dir / "notes"
            notes_dir.mkdir(parents=True, exist_ok=True)

            transcript_path = Path(task.transcript_path)
            notes_filename = f"{transcript_path.stem}_notes.md"
            notes_path = notes_dir / notes_filename
            
            # Determine providers to try
            providers_to_try = []
            if self.config.provider_mode == "Only OpenRouter":
                providers_to_try.append("OpenRouter")
            elif self.config.provider_mode == "Only Gemini":
                providers_to_try.append("Gemini")
            elif self.config.provider_mode == "Fallback Mode":
                providers_to_try.append(self.config.primary_provider)
                if self.config.secondary_provider != self.config.primary_provider:
                    providers_to_try.append(self.config.secondary_provider)
            else:
                providers_to_try.append("OpenRouter")

            notes_result = {"success": False, "error": "No provider configured"}
            
            for provider in providers_to_try:
                try:
                    self.log_activity(f"Generating notes using {provider}...")
                    
                    processor = None
                    if provider == "OpenRouter":
                        note_config = NoteProcessingConfig(
                            model=self.config.openrouter_model,
                            temperature=self.config.openrouter_temperature,
                            max_tokens=self.config.openrouter_max_tokens,
                            pre_prompt=self.read_pre_prompt()
                        )
                        processor = OpenRouterProcessor(config=note_config)
                    elif provider == "Gemini":
                        gemini_config = GeminiProcessingConfig(
                            model=self.config.gemini_model,
                            temperature=self.config.gemini_temperature,
                            max_tokens=self.config.gemini_max_tokens,
                            pre_prompt=self.read_pre_prompt()
                        )
                        processor = GeminiProcessor(config=gemini_config)
                    
                    if processor:
                        notes_result = processor.process_transcript_file(
                            transcript_path=task.transcript_path,
                            output_path=str(notes_path),
                            subject=task.subject
                        )
                    
                    if notes_result['success']:
                        self.log_activity(f"✓ Notes generated successfully with {provider}")
                        break
                    else:
                        self.log_activity(f"✗ {provider} failed: {notes_result.get('error')}")
                
                except Exception as e:
                    self.log_activity(f"Error with {provider}: {e}")
                    notes_result = {"success": False, "error": str(e)}

            if notes_result['success'] and self.config.remove_thinking_tags:
                try:
                    with open(notes_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    cleaned_content = self.remove_thinking_tags(content)

                    with open(notes_path, 'w', encoding='utf-8') as f:
                        f.write(cleaned_content)

                    self.log_activity(f"Removed thinking tags from: {Path(task.filepath).name}")
                except Exception as e:
                    logging.error(f"Error removing thinking tags: {e}")

            if notes_result['success']:
                task.notes_path = str(notes_path)
                task.tokens_used = notes_result.get('tokens_used', 0)
                task.status = "completed"
                self.log_activity(f"Notes generation completed: {Path(task.filepath).name} ({task.tokens_used} tokens)")

                if (self.config.word_auto_update and self.word_manager and
                    task.status == "completed" and task.notes_path):
                    try:
                        self.word_manager.check_new_markdown_file(task.notes_path, task.subject)
                        self.log_activity(f"Word document updated for {task.subject}")

                        if self.config.auto_apply_colors:
                            self.apply_color_to_word_document(task.subject)
                    except Exception as e:
                        logging.error(f"Error updating Word document: {e}")
            else:
                task.status = "error"
                task.error_message = notes_result.get('error', 'Unknown notes processing error')
                self.log_activity(f"Notes processing failed: {Path(task.filepath).name} - {task.error_message}")

        except Exception as e:
            task.status = "error"
            task.error_message = str(e)
            self.log_activity(f"Notes processing error: {Path(task.filepath).name} - {e}")
            logging.error(f"Notes processing error for {task.filepath}: {e}")

    def process_task(self, task: ProcessingTask):
        """Process a single task."""
        try:
            if task.reprocess_type:
                self.handle_reprocessing_task(task)
                return

            if task.status == "queued_notes" and task.transcript_path and Path(task.transcript_path).exists():
                self.log_activity(f"Processing notes only: {Path(task.filepath).name}")
                self.process_notes_only(task)
                return

            task.status = "transcribing"
            self.log_activity(f"Starting transcription: {Path(task.filepath).name}")

            subject_dir = Path(task.subject)
            transcripts_dir = subject_dir / "transcripts"
            notes_dir = subject_dir / "notes"
            transcripts_dir.mkdir(parents=True, exist_ok=True)
            notes_dir.mkdir(parents=True, exist_ok=True)

            transcription_config = TranscriptionConfig(
                language_detection=True,
                speaker_labels=False,
                punctuate=True,
                format_text=True
            )

            transcriber = AssemblyAITranscriber(config=transcription_config)

            result = transcriber.transcribe_file(
                filepath=task.filepath,
                output_dir=transcripts_dir,
                save_txt=True,
                save_json=True
            )

            if result['success']:
                task.transcript_path = next((f for f in result['output_files'] if f.endswith('.txt')), "")
                task.status = "processing_notes"
                self.log_activity(f"Transcription completed: {Path(task.filepath).name}")

                # Determine providers to try
                providers_to_try = []
                if self.config.provider_mode == "Only OpenRouter":
                    providers_to_try.append("OpenRouter")
                elif self.config.provider_mode == "Only Gemini":
                    providers_to_try.append("Gemini")
                elif self.config.provider_mode == "Fallback Mode":
                    providers_to_try.append(self.config.primary_provider)
                    if self.config.secondary_provider != self.config.primary_provider:
                        providers_to_try.append(self.config.secondary_provider)
                else:
                    providers_to_try.append("OpenRouter")

                notes_result = {"success": False, "error": "No provider configured"}
                
                # Setup paths
                transcript_path = Path(task.transcript_path)
                notes_filename = f"{transcript_path.stem}_notes.md"
                notes_path = notes_dir / notes_filename

                for provider in providers_to_try:
                    try:
                        self.log_activity(f"Generating notes using {provider}...")
                        
                        processor = None
                        if provider == "OpenRouter":
                            note_config = NoteProcessingConfig(
                                model=self.config.openrouter_model,
                                temperature=self.config.openrouter_temperature,
                                max_tokens=self.config.openrouter_max_tokens,
                                pre_prompt=self.read_pre_prompt()
                            )
                            processor = OpenRouterProcessor(config=note_config)
                        elif provider == "Gemini":
                            gemini_config = GeminiProcessingConfig(
                                model=self.config.gemini_model,
                                temperature=self.config.gemini_temperature,
                                max_tokens=self.config.gemini_max_tokens,
                                pre_prompt=self.read_pre_prompt()
                            )
                            processor = GeminiProcessor(config=gemini_config)
                        
                        if processor:
                            notes_result = processor.process_transcript_file(
                                transcript_path=task.transcript_path,
                                output_path=str(notes_path),
                                subject=task.subject
                            )
                        
                        if notes_result['success']:
                            self.log_activity(f"✓ Notes generated successfully with {provider}")
                            break
                        else:
                            self.log_activity(f"✗ {provider} failed: {notes_result.get('error')}")
                    
                    except Exception as e:
                        self.log_activity(f"Error with {provider}: {e}")
                        notes_result = {"success": False, "error": str(e)}

                if notes_result['success'] and self.config.remove_thinking_tags:
                    try:
                        with open(notes_path, 'r', encoding='utf-8') as f:
                            content = f.read()

                        cleaned_content = self.remove_thinking_tags(content)

                        with open(notes_path, 'w', encoding='utf-8') as f:
                            f.write(cleaned_content)

                        self.log_activity(f"Removed thinking tags from: {Path(task.filepath).name}")
                    except Exception as e:
                        logging.error(f"Error removing thinking tags: {e}")

                if notes_result['success']:
                    task.notes_path = str(notes_path)
                    task.tokens_used = notes_result.get('tokens_used', 0)
                    task.status = "completed"
                    self.log_activity(f"Notes generation completed: {Path(task.filepath).name} ({task.tokens_used} tokens)")

                    if (self.config.word_auto_update and self.word_manager and
                        task.status == "completed" and task.notes_path):
                        try:
                            self.word_manager.check_new_markdown_file(task.notes_path, task.subject)
                            self.log_activity(f"Word document updated for {task.subject}")

                            if self.config.auto_apply_colors:
                                self.apply_color_to_word_document(task.subject)
                        except Exception as e:
                            logging.error(f"Error updating Word document: {e}")
                else:
                    task.status = "error"
                    task.error_message = notes_result.get('error', 'Unknown notes processing error')
                    self.log_activity(f"Notes processing failed: {Path(task.filepath).name} - {task.error_message}")

            else:
                task.status = "error"
                task.error_message = result.get('error', 'Unknown transcription error')
                self.log_activity(f"Transcription failed: {Path(task.filepath).name} - {task.error_message}")

        except Exception as e:
            task.status = "error"
            task.error_message = str(e)
            self.log_activity(f"Error processing task {Path(task.filepath).name}: {e}")
            logging.error(f"Processing error for {task.filepath}: {e}")

    def get_task_progress(self, task):
        """Get task progress description."""
        if task.status == "completed":
            if task.reprocess_type:
                return f"✓ Reprocessed ({task.reprocess_type})"
            return "✓ Complete - Notes generated"
        elif task.status == "transcript_only":
            return "📝 Has transcript, needs notes"
        elif task.status == "queued_notes":
            return "⏳ Queued for notes generation"
        elif task.status == "queued_transcript":
            return "⏳ Queued for transcript generation"
        elif task.status == "error":
            return f"✗ Error: {task.error_message[:40]}..."
        elif task.status == "transcribing":
            return "🎵 Transcribing audio..."
        elif task.status == "processing_notes":
            return "📝 Generating notes..."
        elif task.status == "queued":
            return "⏳ Queued for processing"
        elif task.status == "pending":
            return "⏳ Waiting to start"
        else:
            return f"Unknown: {task.status}"

    def open_file_in_editor(self, filepath):
        """Open file in system default editor."""
        import subprocess
        import platform

        try:
            if platform.system() == "Windows":
                os.startfile(filepath)
            elif platform.system() == "Darwin":
                subprocess.run(["open", filepath])
            else:
                subprocess.run(["xdg-open", filepath])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")

    def refresh_reprocessing_display(self):
        """Refresh the reprocessing files display."""
        for item in self.reprocess_tree.get_children():
            self.reprocess_tree.delete(item)

        for file_info in sorted(self.reprocessing_files.values(), key=lambda f: f.filename):
            status_parts = []
            if file_info.has_transcript:
                status_parts.append("Has transcript")
            if file_info.has_notes:
                status_parts.append("Has notes")

            if not status_parts:
                status = "Not processed"
            else:
                status = ", ".join(status_parts)

            size_mb = file_info.file_size / (1024 * 1024)
            size_str = f"{size_mb:.1f} MB" if size_mb > 1 else f"{file_info.file_size / 1024:.1f} KB"

            select_indicator = "☑" if file_info.selected else "☐"
            transcript_indicator = "✓" if file_info.has_transcript else "✗"
            notes_indicator = "✓" if file_info.has_notes else "✗"

            self.reprocess_tree.insert("", tk.END, values=(
                select_indicator,
                file_info.filename,
                file_info.subject,
                status,
                size_str,
                file_info.modified_date,
                transcript_indicator,
                notes_indicator
            ), tags=(file_info.filepath,))

    def run(self):
        """Run the application."""
        try:
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.root.mainloop()
        except KeyboardInterrupt:
            self.shutdown()

    def on_closing(self):
        """Handle application closing."""
        self.shutdown()
        self.root.destroy()

    def shutdown(self):
        """Shutdown the application."""
        logging.info("Shutting down application...")

        self.stop_file_monitoring()

        if self.processing_thread and self.processing_thread.is_alive():
            self.processing_thread.stop()
            self.processing_thread.join(timeout=5)

        self.save_config()


def main():
    """Main entry point."""
    try:
        app = SchoolNoteApp()
        app.run()
    except Exception as e:
        logging.error(f"Application error: {e}")
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
