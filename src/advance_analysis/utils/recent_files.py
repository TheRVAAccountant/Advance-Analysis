"""
Recent Files Manager for the Advance Analysis Tool.

This module manages a history of recently used files for each file type
in the application, allowing users to quickly select from previously used files.
"""
import os
import json
from typing import List, Dict, Optional
from datetime import datetime
import logging

logger = logging.getLogger(__name__)


class RecentFilesManager:
    """Manages recent files history for the application."""
    
    MAX_RECENT_FILES = 10
    
    def __init__(self, config_dir: Optional[str] = None):
        """
        Initialize the Recent Files Manager.
        
        Args:
            config_dir: Directory to store the recent files config.
                       Defaults to user's home directory.
        """
        if config_dir is None:
            config_dir = os.path.expanduser("~")
        
        self.config_file = os.path.join(config_dir, ".advance_analysis_recent_files.json")
        self.recent_files = self._load_recent_files()
    
    def _load_recent_files(self) -> Dict[str, List[Dict[str, str]]]:
        """Load recent files from the config file."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    data = json.load(f)
                    # Ensure all file types exist
                    for file_type in ["advance_analysis", "current_dhstier", "prior_dhstier"]:
                        if file_type not in data:
                            data[file_type] = []
                    return data
        except Exception as e:
            logger.warning(f"Could not load recent files: {e}")
        
        return {
            "advance_analysis": [],
            "current_dhstier": [],
            "prior_dhstier": []
        }
    
    def _save_recent_files(self) -> None:
        """Save recent files to the config file."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.recent_files, f, indent=2)
        except Exception as e:
            logger.error(f"Could not save recent files: {e}")
    
    def add_file(self, file_type: str, file_path: str) -> None:
        """
        Add a file to the recent files list.
        
        Args:
            file_type: Type of file ("advance_analysis", "current_dhstier", "prior_dhstier")
            file_path: Path to the file
        """
        if file_type not in self.recent_files:
            logger.warning(f"Unknown file type: {file_type}")
            return
        
        # Normalize the path
        file_path = os.path.abspath(file_path)
        
        # Check if file exists
        if not os.path.exists(file_path):
            logger.warning(f"File does not exist: {file_path}")
            return
        
        # Get file info
        file_info = {
            "path": file_path,
            "name": os.path.basename(file_path),
            "directory": os.path.dirname(file_path),
            "last_used": datetime.now().isoformat(),
            "size": os.path.getsize(file_path)
        }
        
        # Remove if already in list
        self.recent_files[file_type] = [
            f for f in self.recent_files[file_type] 
            if f.get("path") != file_path
        ]
        
        # Add to beginning of list
        self.recent_files[file_type].insert(0, file_info)
        
        # Trim to max size
        self.recent_files[file_type] = self.recent_files[file_type][:self.MAX_RECENT_FILES]
        
        # Save changes
        self._save_recent_files()
    
    def get_recent_files(self, file_type: str) -> List[Dict[str, str]]:
        """
        Get list of recent files for a specific type.
        
        Args:
            file_type: Type of file
            
        Returns:
            List of file info dictionaries
        """
        # Filter out files that no longer exist
        valid_files = []
        for file_info in self.recent_files.get(file_type, []):
            if os.path.exists(file_info.get("path", "")):
                valid_files.append(file_info)
        
        # Update the list if files were removed
        if len(valid_files) != len(self.recent_files.get(file_type, [])):
            self.recent_files[file_type] = valid_files
            self._save_recent_files()
        
        return valid_files
    
    def clear_recent_files(self, file_type: Optional[str] = None) -> None:
        """
        Clear recent files.
        
        Args:
            file_type: Specific file type to clear, or None to clear all
        """
        if file_type:
            if file_type in self.recent_files:
                self.recent_files[file_type] = []
        else:
            for key in self.recent_files:
                self.recent_files[key] = []
        
        self._save_recent_files()
    
    def format_file_display(self, file_info: Dict[str, str]) -> str:
        """
        Format file info for display.
        
        Args:
            file_info: File information dictionary
            
        Returns:
            Formatted string for display
        """
        name = file_info.get("name", "Unknown")
        directory = file_info.get("directory", "")
        
        # Truncate long paths
        if len(directory) > 50:
            parts = directory.split(os.sep)
            if len(parts) > 4:
                directory = os.sep.join(parts[:2] + ["..."] + parts[-2:])
        
        # Format last used time
        last_used = file_info.get("last_used", "")
        if last_used:
            try:
                dt = datetime.fromisoformat(last_used)
                days_ago = (datetime.now() - dt).days
                if days_ago == 0:
                    time_str = "Today"
                elif days_ago == 1:
                    time_str = "Yesterday"
                else:
                    time_str = f"{days_ago} days ago"
            except:
                time_str = ""
        else:
            time_str = ""
        
        # Build display string
        display_parts = [name]
        if directory:
            display_parts.append(f"({directory})")
        if time_str:
            display_parts.append(f"- {time_str}")
        
        return " ".join(display_parts)