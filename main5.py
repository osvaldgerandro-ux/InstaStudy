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
from dataclasses import dataclass, asdict
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from word_document_manager import WordDocumentManager, WordFormattingConfig

try:
    from transcriber import AssemblyAITranscriber, TranscriptionConfig
    from openrouter_processor import OpenRouterProcessor, NoteProcessingConfig
    from color_changer import COLOR_PALETTES, change_theme_colors
except ImportError as e:
    print(f"Error: Could not import required modules: {e}")
    print("Make sure transcriber.py, openrouter_processor.py, word_document_manager.py, and color_changer.py are in the same directory.")
    sys.exit(1)


@dataclass
class AppConfig:
    """Application configuration."""
    watch_directory: str = ""
    subjects: List[str] = None
    auto_process: bool = True
    supported_extensions: List[str] = None
    openrouter_model: str = "openai/gpt-4o-mini"
    openrouter_temperature: float = 0.3
    openrouter_max_tokens: int = 4000
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
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.config = AppConfig(**data)
                logging.info("Configuration loaded successfully")
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

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.setup_config_tab(notebook)
        self.setup_api_keys_tab(notebook)
        self.setup_pre_prompt_tab(notebook)
        self.setup_monitoring_tab(notebook)
        self.setup_tasks_tab(notebook)
        self.setup_reprocessing_tab(notebook)
        self.setup_word_tab(notebook)
        self.setup_color_management_tab(notebook)

        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.update_gui()

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

        ttk.Label(config_frame, text="Watch Directory:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.watch_dir_var = tk.StringVar(value=self.config.watch_directory)
        ttk.Entry(config_frame, textvariable=self.watch_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_watch_directory).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(config_frame, text="OpenRouter Model:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.model_var = tk.StringVar(value=self.config.openrouter_model)
        model_combo = ttk.Combobox(config_frame, textvariable=self.model_var, width=47)
        model_combo.grid(row=1, column=1, padx=5, pady=5)

        model_combo['values'] = [
            "openai/gpt-4o-mini",
            "openai/gpt-4o",
            "anthropic/claude-3-haiku",
            "anthropic/claude-3-sonnet",
            "meta-llama/llama-3.1-8b-instruct",
            "google/gemini-pro-1.5"
        ]

        ttk.Button(config_frame, text="List Models", command=self.list_available_models).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(config_frame, text="Temperature:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.temperature_var = tk.DoubleVar(value=self.config.openrouter_temperature)
        ttk.Scale(config_frame, from_=0.0, to=2.0, variable=self.temperature_var,
                 orient=tk.HORIZONTAL, length=300).grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        self.temp_label = ttk.Label(config_frame, text=f"{self.config.openrouter_temperature:.1f}")
        self.temp_label.grid(row=2, column=2, padx=5, pady=5)

        self.temperature_var.trace('w', self.update_temperature_label)

        ttk.Label(config_frame, text="Max Tokens:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=5)
        self.max_tokens_var = tk.IntVar(value=self.config.openrouter_max_tokens)
        ttk.Spinbox(config_frame, from_=1000, to=8000, textvariable=self.max_tokens_var,
                   width=48).grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(config_frame, text="Subjects:").grid(row=4, column=0, sticky=tk.NW, padx=5, pady=5)

        subjects_frame = ttk.Frame(config_frame)
        subjects_frame.grid(row=4, column=1, columnspan=2, sticky=tk.W, padx=5, pady=5)

        self.subjects_listbox = tk.Listbox(subjects_frame, height=6, width=30)
        self.subjects_listbox.pack(side=tk.LEFT, padx=(0, 5))

        self.update_subjects_listbox()

        subjects_controls = ttk.Frame(subjects_frame)
        subjects_controls.pack(side=tk.LEFT, fill=tk.Y)

        self.subject_entry = ttk.Entry(subjects_controls, width=20)
        self.subject_entry.pack(pady=(0, 5))

        ttk.Button(subjects_controls, text="Add Subject", command=self.add_subject).pack(fill=tk.X, pady=(0, 2))
        ttk.Button(subjects_controls, text="Remove Selected", command=self.remove_subject).pack(fill=tk.X, pady=(0, 2))

        self.auto_process_var = tk.BooleanVar(value=self.config.auto_process)
        ttk.Checkbutton(config_frame, text="Auto-process new files", variable=self.auto_process_var).grid(row=5, column=1, sticky=tk.W, padx=5, pady=5)

        self.remove_thinking_var = tk.BooleanVar(value=self.config.remove_thinking_tags)
        ttk.Checkbutton(config_frame, text="Remove <think> tags from generated notes",
                        variable=self.remove_thinking_var).grid(row=6, column=1, sticky=tk.W, padx=5, pady=5)

        ttk.Button(config_frame, text="Save Configuration", command=self.save_configuration).grid(row=7, column=1, pady=20)

    def setup_api_keys_tab(self, notebook):
        """Setup API keys management tab."""
        api_frame = ttk.Frame(notebook)
        notebook.add(api_frame, text="API Keys")

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

        self.load_api_key(self.assemblyai_key_file, self.assemblyai_key_var)
        self.load_api_key(self.openrouter_key_file, self.openrouter_key_var)

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

            scan_path = Path(scan_dir)
            found_count = 0

            for file_path in scan_path.rglob("*"):
                if file_path.is_file() and file_path.suffix.lower() in self.config.supported_extensions:
                    filename = file_path.name.lower()
                    matching_subject = None

                    for subject in self.config.subjects:
                        if subject.lower() in filename:
                            matching_subject = subject
                            break

                    if not matching_subject:
                        continue

                    file_info = ReprocessingFileInfo(
                        filepath=str(file_path),
                        filename=file_path.name,
                        subject=matching_subject,
                        file_size=file_path.stat().st_size,
                        modified_date=datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                    )

                    subject_dir = Path(matching_subject)
                    transcripts_dir = subject_dir / "transcripts"
                    notes_dir = subject_dir / "notes"

                    transcript_path = transcripts_dir / f"{file_path.stem}.txt"
                    notes_path = notes_dir / f"{file_path.stem}_notes.md"

                    if transcript_path.exists():
                        file_info.has_transcript = True
                        file_info.transcript_path = str(transcript_path)

                    if notes_path.exists():
                        file_info.has_notes = True
                        file_info.notes_path = str(notes_path)

                    self.reprocessing_files[str(file_path)] = file_info
                    found_count += 1

            self.refresh_reprocessing_display()
            self.reprocess_status_var.set(f"Found {found_count} audio files")
            self.log_activity(f"Scanned {scan_dir}: found {found_count} processable audio files")

        except Exception as e:
            self.reprocess_status_var.set(f"Error scanning directory: {e}")
            messagebox.showerror("Error", f"Failed to scan directory: {e}")
            logging.error(f"Reprocessing scan error: {e}")

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
                size_str,
                file_info.modified_date,
                transcript_indicator,
                notes_indicator,
                status
            ))

        selected_count = len(self.selected_files)
        total_count = len(self.reprocessing_files)
        self.selection_count_var.set(f"{selected_count} of {total_count} files selected")

    def on_reprocess_tree_click(self, event):
        """Handle click on reprocessing tree."""
        region = self.reprocess_tree.identify_region(event.x, event.y)
        if region == "cell":
            item = self.reprocess_tree.identify_row(event.y)
            column = self.reprocess_tree.identify_column(event.x)

            if column == "#1" and item:
                values = self.reprocess_tree.item(item, "values")
                if len(values) > 1:
                    filename = values[1]

                    for filepath, file_info in self.reprocessing_files.items():
                        if file_info.filename == filename:
                            file_info.selected = not file_info.selected
                            if file_info.selected:
                                self.selected_files.add(filepath)
                            else:
                                self.selected_files.discard(filepath)
                            break

                    self.refresh_reprocessing_display()

    def on_reprocess_tree_double_click(self, event):
        """Handle double-click on reprocessing tree."""
        item = self.reprocess_tree.selection()[0] if self.reprocess_tree.selection() else None
        if item:
            values = self.reprocess_tree.item(item, "values")
            if len(values) > 1:
                filename = values[1]

                for file_info in self.reprocessing_files.values():
                    if file_info.filename == filename:
                        self.show_file_details(file_info)
                        break

    def show_file_details(self, file_info: ReprocessingFileInfo):
        """Show detailed information about a file."""
        details_window = tk.Toplevel(self.root)
        details_window.title(f"File Details - {file_info.filename}")
        details_window.geometry("600x400")
        details_window.transient(self.root)

        details_text = scrolledtext.ScrolledText(details_window, wrap=tk.WORD, font=("Consolas", 10))
        details_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        details = f"""File: {file_info.filename}
Path: {file_info.filepath}
Subject: {file_info.subject}
Size: {file_info.file_size / (1024*1024):.2f} MB
Modified: {file_info.modified_date}
Has Transcript: {'Yes' if file_info.has_transcript else 'No'}
Has Notes: {'Yes' if file_info.has_notes else 'No'}

Transcript Path: {file_info.transcript_path if file_info.has_transcript else 'Not found'}
Notes Path: {file_info.notes_path if file_info.has_notes else 'Not found'}
"""

        details_text.insert(tk.END, details)
        details_text.config(state=tk.DISABLED)

        buttons_frame = ttk.Frame(details_window)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(buttons_frame, text="Close", command=details_window.destroy).pack(side=tk.RIGHT)

        if file_info.has_transcript:
            ttk.Button(buttons_frame, text="Open Transcript",
                      command=lambda: self.open_file_in_editor(file_info.transcript_path)).pack(side=tk.LEFT, padx=(0, 5))

        if file_info.has_notes:
            ttk.Button(buttons_frame, text="Open Notes",
                      command=lambda: self.open_file_in_editor(file_info.notes_path)).pack(side=tk.LEFT, padx=(0, 5))

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

    def select_all_files(self):
        """Select all files."""
        for filepath, file_info in self.reprocessing_files.items():
            file_info.selected = True
            self.selected_files.add(filepath)
        self.refresh_reprocessing_display()

    def deselect_all_files(self):
        """Deselect all files."""
        for file_info in self.reprocessing_files.values():
            file_info.selected = False
        self.selected_files.clear()
        self.refresh_reprocessing_display()

    def select_by_subject(self):
        """Select files by subject."""
        if not self.config.subjects:
            messagebox.showwarning("Warning", "No subjects configured!")
            return

        subject_window = tk.Toplevel(self.root)
        subject_window.title("Select by Subject")
        subject_window.geometry("300x200")
        subject_window.transient(self.root)

        ttk.Label(subject_window, text="Select subject:").pack(pady=10)

        subject_var = tk.StringVar(value=self.config.subjects[0])
        subject_combo = ttk.Combobox(subject_window, textvariable=subject_var,
                                   values=self.config.subjects, state="readonly")
        subject_combo.pack(pady=10)

        def apply_selection():
            selected_subject = subject_var.get()
            count = 0
            for filepath, file_info in self.reprocessing_files.items():
                if file_info.subject == selected_subject:
                    file_info.selected = True
                    self.selected_files.add(filepath)
                    count += 1

            self.refresh_reprocessing_display()
            subject_window.destroy()
            messagebox.showinfo("Selection", f"Selected {count} files for subject '{selected_subject}'")

        ttk.Button(subject_window, text="Select Files", command=apply_selection).pack(pady=10)
        ttk.Button(subject_window, text="Cancel", command=subject_window.destroy).pack(pady=5)

    def reprocess_selected_files(self):
        """Reprocess selected files."""
        if not self.selected_files:
            messagebox.showwarning("Warning", "No files selected for reprocessing!")
            return

        if not self.read_api_key_file(self.assemblyai_key_file):
            messagebox.showerror("Error", "AssemblyAI API key not found. Please configure it first!")
            return

        if not self.read_api_key_file(self.openrouter_key_file):
            messagebox.showerror("Error", "OpenRouter API key not found. Please configure it first!")
            return

        reprocess_type = self.reprocess_type_var.get()
        selected_count = len(self.selected_files)

        type_desc = {
            "transcript": "transcripts only",
            "notes": "notes only",
            "both": "both transcripts and notes"
        }

        if not messagebox.askyesno("Confirm Reprocessing",
                                  f"Reprocess {selected_count} files for {type_desc[reprocess_type]}?\n\n"
                                  f"This will overwrite existing files of the selected type(s)."):
            return

        tasks_created = 0
        for filepath in self.selected_files:
            if filepath in self.reprocessing_files:
                file_info = self.reprocessing_files[filepath]

                task = ProcessingTask(
                    filepath=filepath,
                    subject=file_info.subject,
                    reprocess_type=reprocess_type
                )

                if file_info.has_transcript:
                    task.transcript_path = file_info.transcript_path
                if file_info.has_notes:
                    task.notes_path = file_info.notes_path

                if reprocess_type == "notes" and file_info.has_transcript:
                    task.status = "queued_notes"
                elif reprocess_type == "transcript":
                    task.status = "queued_transcript"
                else:
                    task.status = "queued"

                task_key = f"reprocess_{filepath}_{datetime.now().timestamp()}"
                self.tasks[task_key] = task
                self.task_queue.put(task)
                tasks_created += 1

        self.reprocess_status_var.set(f"Queued {tasks_created} files for reprocessing")
        self.log_activity(f"Queued {tasks_created} files for reprocessing ({reprocess_type})")

        messagebox.showinfo("Reprocessing Started",
                           f"Queued {tasks_created} files for reprocessing.\n"
                           f"Check the 'Processing Tasks' tab to monitor progress.")

    def delete_selected_outputs(self):
        """Delete output files for selected files."""
        if not self.selected_files:
            messagebox.showwarning("Warning", "No files selected!")
            return

        files_with_outputs = []
        for filepath in self.selected_files:
            if filepath in self.reprocessing_files:
                file_info = self.reprocessing_files[filepath]
                if file_info.has_transcript or file_info.has_notes:
                    files_with_outputs.append(file_info)

        if not files_with_outputs:
            messagebox.showinfo("Info", "No output files found for selected audio files.")
            return

        if not messagebox.askyesno("Confirm Deletion",
                                  f"Delete transcript and/or note files for {len(files_with_outputs)} selected files?\n\n"
                                  f"This cannot be undone!"):
            return

        deleted_count = 0
        errors = []

        for file_info in files_with_outputs:
            try:
                if file_info.has_transcript and file_info.transcript_path:
                    transcript_path = Path(file_info.transcript_path)
                    if transcript_path.exists():
                        transcript_path.unlink()
                        deleted_count += 1

                        json_path = transcript_path.with_suffix('.json')
                        if json_path.exists():
                            json_path.unlink()

                if file_info.has_notes and file_info.notes_path:
                    notes_path = Path(file_info.notes_path)
                    if notes_path.exists():
                        notes_path.unlink()
                        deleted_count += 1

            except Exception as e:
                errors.append(f"{file_info.filename}: {e}")

        self.scan_reprocess_files()

        result_msg = f"Deleted {deleted_count} output files."
        if errors:
            result_msg += f"\n\nErrors:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                result_msg += f"\n... and {len(errors) - 5} more errors."

        messagebox.showinfo("Deletion Complete", result_msg)
        self.log_activity(f"Deleted {deleted_count} output files for selected audio files")

    def open_selected_file_location(self):
        """Open file location for selected files."""
        if not self.selected_files:
            messagebox.showwarning("Warning", "No files selected!")
            return

        first_file = next(iter(self.selected_files))
        if first_file in self.reprocessing_files:
            file_path = Path(first_file)
            folder_path = file_path.parent

            try:
                import subprocess
                import platform

                if platform.system() == "Windows":
                    subprocess.run(["explorer", "/select,", str(file_path)])
                elif platform.system() == "Darwin":
                    subprocess.run(["open", "-R", str(file_path)])
                else:
                    subprocess.run(["xdg-open", str(folder_path)])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file location: {e}")

    def update_line_spacing_label(self, *args):
        """Update line spacing label."""
        self.word_line_spacing_label.config(text=f"{self.word_line_spacing_var.get():.2f}")

    def save_word_configuration(self):
        """Save Word document configuration."""
        try:
            self.config.word_auto_update = self.word_auto_update_var.get()
            self.config.word_font_name = self.word_font_name_var.get()
            self.config.word_font_size = self.word_font_size_var.get()
            self.config.word_line_spacing = self.word_line_spacing_var.get()

            for i in range(1, 4):
                setattr(self.config, f'word_heading{i}_size', self.word_heading_vars[i].get())

            self.save_config()
            self.init_word_manager()
            messagebox.showinfo("Success", "Word document settings saved!")
            self.log_activity("Word document settings updated")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Word settings: {e}")

    def update_all_word_documents(self):
        """Update all Word documents."""
        if not self.word_manager:
            messagebox.showerror("Error", "Word document manager not initialized!")
            return

        try:
            self.word_status_var.set("Updating Word documents...")
            self.root.update()

            results = self.word_manager.update_all_subjects(self.config.subjects)

            updated_count = sum(1 for result in results.values() if result.get('updated', False))

            if updated_count > 0:
                self.word_status_var.set(f"Updated {updated_count} Word document(s)")
                self.log_activity(f"Updated {updated_count} Word documents")
            else:
                self.word_status_var.set("All Word documents up to date")
                self.log_activity("All Word documents are up to date")

            result_msg = []
            for subject, result in results.items():
                if result.get('updated'):
                    changes = result.get('changes', [])
                    result_msg.append(f"{subject}: Updated ({', '.join(changes[:3])}{'...' if len(changes) > 3 else ''})")
                elif result.get('success'):
                    result_msg.append(f"{subject}: Up to date")
                else:
                    result_msg.append(f"{subject}: Error - {result.get('error', 'Unknown')}")

            if result_msg:
                messagebox.showinfo("Word Documents Update", "\n".join(result_msg))

        except Exception as e:
            self.word_status_var.set("Error updating Word documents")
            messagebox.showerror("Error", f"Failed to update Word documents: {e}")
            logging.error(f"Error updating Word documents: {e}")

    def regenerate_all_word_documents(self):
        """Regenerate all Word documents from scratch."""
        if not self.word_manager:
            messagebox.showerror("Error", "Word document manager not initialized!")
            return

        if not messagebox.askyesno("Confirm Regeneration",
                                  "This will regenerate all Word documents from scratch. Continue?"):
            return

        try:
            self.word_status_var.set("Regenerating all Word documents...")
            self.root.update()

            results = self.word_manager.regenerate_all_documents(self.config.subjects)

            success_count = sum(1 for result in results.values() if result.get('success', False))

            self.word_status_var.set(f"Regenerated {success_count} Word document(s)")
            self.log_activity(f"Regenerated {success_count} Word documents from scratch")

            if self.config.auto_apply_colors:
                self.log_activity("Applying colors to regenerated documents...")
                for subject in self.config.subjects:
                    self.apply_color_to_word_document(subject)

            messagebox.showinfo("Success", f"Successfully regenerated {success_count} Word documents!")

        except Exception as e:
            self.word_status_var.set("Error regenerating Word documents")
            messagebox.showerror("Error", f"Failed to regenerate Word documents: {e}")
            logging.error(f"Error regenerating Word documents: {e}")

    def open_word_documents_folder(self):
        """Open the folder containing Word documents."""
        import subprocess
        import platform
    
        # Search for Word documents in multiple locations
        possible_dirs = [
            Path("Appunti Completi"),
            Path("."),
        ]
        
        # Also check subject-specific directories
        for subject in self.config.subjects:
            possible_dirs.extend([
                Path("Appunti Completi") / subject,
                Path(subject) / "Appunti Completi",
                Path(subject)
            ])
    
        found_folder = None
        for dir_path in possible_dirs:
            if dir_path.exists():
                word_files = list(dir_path.glob("*_combined_notes.docx"))
                if word_files:
                    found_folder = dir_path.resolve()
                    break
    
        if not found_folder:
            messagebox.showinfo("Info", "No Word documents found. Generate some documents first!")
            return
    
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
        self.config.openrouter_model = self.model_var.get()
        self.config.openrouter_temperature = self.temperature_var.get()
        self.config.openrouter_max_tokens = self.max_tokens_var.get()
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

        if not self.read_api_key_file(self.openrouter_key_file):
            messagebox.showerror("Error", "OpenRouter API key not found. Please configure it first!")
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

        except Exception as e:
            messagebox.showerror("Error", f"Failed to start monitoring: {e}")
            logging.error(f"Failed to start monitoring: {e}")

    def stop_file_monitoring(self):
        """Stop file monitoring."""
        if self.observer and self.observer.is_alive():
            self.observer.stop()
            self.observer.join()
            self.log_activity("Stopped file monitoring")
            logging.info("Stopped file monitoring")

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

            note_config = NoteProcessingConfig(
                model=self.config.openrouter_model,
                temperature=self.config.openrouter_temperature,
                max_tokens=self.config.openrouter_max_tokens,
                pre_prompt=self.read_pre_prompt()
            )

            processor = OpenRouterProcessor(config=note_config)

            # Notes processing delegated to process_notes_only


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

                note_config = NoteProcessingConfig(
                    model=self.config.openrouter_model,
                    temperature=self.config.openrouter_temperature,
                    max_tokens=self.config.openrouter_max_tokens,
                    pre_prompt=self.read_pre_prompt()
                )

                processor = OpenRouterProcessor(config=note_config)

                transcript_path = Path(task.transcript_path)
                notes_filename = f"{transcript_path.stem}_notes.md"
                notes_path = notes_dir / notes_filename

                notes_result = processor.process_transcript_file(
                    transcript_path=task.transcript_path,
                    output_path=str(notes_path),
                    subject=task.subject
                )

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
                    self.log_activity(f"Notes processing completed: {Path(task.filepath).name} ({task.tokens_used} tokens)")

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
            self.log_activity(f"Processing error: {Path(task.filepath).name} - {e}")
            logging.error(f"Processing error for {task.filepath}: {e}")

    def handle_reprocessing_task(self, task: ProcessingTask):
        """Handle reprocessing tasks."""
        try:
            reprocess_type = task.reprocess_type
            self.log_activity(f"Reprocessing {reprocess_type}: {Path(task.filepath).name}")

            subject_dir = Path(task.subject)
            transcripts_dir = subject_dir / "transcripts"
            notes_dir = subject_dir / "notes"
            transcripts_dir.mkdir(parents=True, exist_ok=True)
            notes_dir.mkdir(parents=True, exist_ok=True)

            if reprocess_type in ["transcript", "both"]:
                task.status = "transcribing"

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
                    self.log_activity(f"Transcript reprocessed: {Path(task.filepath).name}")
                else:
                    task.status = "error"
                    task.error_message = result.get('error', 'Transcription failed')
                    return

            if reprocess_type in ["notes", "both"] and task.transcript_path:
                task.status = "processing_notes"

                note_config = NoteProcessingConfig(
                    model=self.config.openrouter_model,
                    temperature=self.config.openrouter_temperature,
                    max_tokens=self.config.openrouter_max_tokens,
                    pre_prompt=self.read_pre_prompt()
                )

                processor = OpenRouterProcessor(config=note_config)

                transcript_path = Path(task.transcript_path)
                notes_filename = f"{transcript_path.stem}_notes.md"
                notes_path = notes_dir / notes_filename

                notes_result = processor.process_transcript_file(
                    transcript_path=task.transcript_path,
                    output_path=str(notes_path),
                    subject=task.subject
                )

                if notes_result['success']:
                    task.notes_path = str(notes_path)
                    task.tokens_used = notes_result.get('tokens_used', 0)
                    self.log_activity(f"Notes reprocessed: {Path(task.filepath).name}")

                    if self.config.remove_thinking_tags:
                        try:
                            with open(notes_path, 'r', encoding='utf-8') as f:
                                content = f.read()

                            cleaned_content = self.remove_thinking_tags(content)

                            with open(notes_path, 'w', encoding='utf-8') as f:
                                f.write(cleaned_content)
                        except Exception as e:
                            logging.error(f"Error removing thinking tags: {e}")
                else:
                    task.status = "error"
                    task.error_message = notes_result.get('error', 'Notes processing failed')
                    return

            task.status = "completed"
            self.log_activity(f"Reprocessing completed: {Path(task.filepath).name}")

            if (self.config.word_auto_update and self.word_manager and
                task.notes_path and reprocess_type in ["notes", "both"]):
                try:
                    self.word_manager.check_new_markdown_file(task.notes_path, task.subject)
                    self.log_activity(f"Word document updated for {task.subject}")

                    if self.config.auto_apply_colors:
                        self.apply_color_to_word_document(task.subject)
                except Exception as e:
                    logging.error(f"Error updating Word document: {e}")

        except Exception as e:
            task.status = "error"
            task.error_message = str(e)
            self.log_activity(f"Reprocessing error: {Path(task.filepath).name} - {e}")
            logging.error(f"Reprocessing error for {task.filepath}: {e}")

    def log_activity(self, message):
        """Log activity to the GUI."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"

        def update_gui():
            self.activity_text.config(state=tk.NORMAL)
            self.activity_text.insert(tk.END, full_message)
            self.activity_text.see(tk.END)
            self.activity_text.config(state=tk.DISABLED)

        if self.root:
            self.root.after(0, update_gui)

    def update_gui(self):
        """Update GUI elements periodically."""
        try:
            if self.observer and self.observer.is_alive():
                self.monitor_status_var.set(f"✓ Monitoring: {self.config.watch_directory}")
            else:
                self.monitor_status_var.set("✗ Not monitoring")

            active_tasks = sum(1 for task in self.tasks.values() if task.status in ["queued", "transcribing", "processing_notes", "queued_notes", "queued_transcript"])
            completed_tasks = sum(1 for task in self.tasks.values() if task.status == "completed")
            total_tokens = sum(task.tokens_used for task in self.tasks.values() if task.tokens_used > 0)

            if active_tasks > 0:
                self.status_var.set(f"Processing {active_tasks} task(s) | {completed_tasks} completed | {total_tokens} tokens used")
            else:
                self.status_var.set(f"Ready | {completed_tasks} completed | {total_tokens} tokens used")

            self.refresh_tasks_display()

        except Exception as e:
            logging.error(f"Error updating GUI: {e}")

        if self.root:
            self.root.after(2000, self.update_gui)

    def refresh_tasks_display(self):
        """Refresh the tasks display."""
        for item in self.tasks_tree.get_children():
            self.tasks_tree.delete(item)

        for task in sorted(self.tasks.values(), key=lambda t: t.created_at, reverse=True):
            filename = Path(task.filepath).name
            progress = self.get_task_progress(task)
            tokens = str(task.tokens_used) if task.tokens_used > 0 else ""

            self.tasks_tree.insert("", 0, values=(
                filename,
                task.subject,
                task.status,
                task.created_at,
                tokens,
                progress
            ))

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
        else:
            return "⏸ Pending"

    def clear_completed_tasks(self):
        """Clear completed tasks."""
        completed_tasks = [path for path, task in self.tasks.items() if task.status == "completed"]
        for path in completed_tasks:
            del self.tasks[path]
        self.refresh_tasks_display()
        self.log_activity(f"Cleared {len(completed_tasks)} completed tasks")

    def retry_failed_tasks(self):
        """Retry failed tasks."""
        failed_tasks = [task for task in self.tasks.values() if task.status == "error"]
        for task in failed_tasks:
            task.status = "pending"
            task.error_message = ""
            self.task_queue.put(task)
        self.log_activity(f"Retrying {len(failed_tasks)} failed tasks")

    def open_notes_folder(self):
        """Open the notes folder in file explorer."""
        import subprocess
        import platform

        notes_folder = None
        for subject in self.config.subjects:
            subject_notes_dir = Path(subject) / "notes"
            if subject_notes_dir.exists():
                notes_folder = subject_notes_dir
                break

        if not notes_folder:
            if self.config.subjects:
                notes_folder = Path(self.config.subjects[0]) / "notes"
                notes_folder.mkdir(parents=True, exist_ok=True)

        if notes_folder and notes_folder.exists():
            try:
                if platform.system() == "Windows":
                    os.startfile(str(notes_folder))
                elif platform.system() == "Darwin":
                    subprocess.run(["open", str(notes_folder)])
                else:
                    subprocess.run(["xdg-open", str(notes_folder)])
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder: {e}")
        else:
            messagebox.showinfo("Info", "No notes folder found. Process some files first!")

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
