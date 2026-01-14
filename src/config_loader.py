"""
Configuration loader for FIFRA Automation.
Loads configuration from YAML file and environment variables.
"""

import os
import yaml
from pathlib import Path
from typing import Dict, Any


class Config:
    """Configuration manager for the application."""
    
    def __init__(self, config_path: str = None):
        """
        Initialize configuration.
        
        Args:
            config_path: Path to config.yaml file. If None, uses default location.
        """
        if config_path is None:
            # Default to config/config.yaml relative to project root
            project_root = Path(__file__).parent.parent
            config_path = project_root / "config" / "config.yaml"
        
        self.config_path = Path(config_path)
        self._config = self._load_config()
        self._override_with_env_vars()
    
    def _load_config(self) -> Dict[str, Any]:
        """Load configuration from YAML file."""
        if not self.config_path.exists():
            raise FileNotFoundError(f"Configuration file not found: {self.config_path}")
        
        with open(self.config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
        
        return config
    
    def _override_with_env_vars(self):
        """Override configuration values with environment variables."""
        # Override credentials if set in environment
        if os.getenv('ENLABEL_USERNAME'):
            self._config['enlabel']['username'] = os.getenv('ENLABEL_USERNAME')
        if os.getenv('ENLABEL_PASSWORD'):
            self._config['enlabel']['password'] = os.getenv('ENLABEL_PASSWORD')
    
    def get(self, key_path: str, default: Any = None) -> Any:
        """
        Get configuration value using dot notation.
        
        Args:
            key_path: Dot-separated path to config value (e.g., 'enlabel.login_url')
            default: Default value if key not found
        
        Returns:
            Configuration value or default
        """
        keys = key_path.split('.')
        value = self._config
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        
        return value
    
    def get_section(self, section: str) -> Dict[str, Any]:
        """Get a configuration section."""
        return self._config.get(section, {})
    
    @property
    def enlabel_username(self) -> str:
        """Get Enlabel username."""
        return self.get('enlabel.username', '')
    
    @property
    def enlabel_password(self) -> str:
        """Get Enlabel password."""
        return self.get('enlabel.password', '')
    
    @property
    def enlabel_login_url(self) -> str:
        """Get Enlabel login URL."""
        return self.get('enlabel.login_url', '')
    
    @property
    def enlabel_manage_databases_url(self) -> str:
        """Get Enlabel ManageDatabases URL."""
        return self.get('enlabel.manage_databases_url', '')


# Global configuration instance
_config_instance: Config = None


def get_config(config_path: str = None) -> Config:
    """
    Get global configuration instance (singleton pattern).
    
    Args:
        config_path: Path to config file (only used on first call)
    
    Returns:
        Config instance
    """
    global _config_instance
    if _config_instance is None:
        _config_instance = Config(config_path)
    return _config_instance
