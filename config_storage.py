"""
Config Storage Module
Handles persistence of Zabbix connections and item templates in JSON format.
"""

import os
import json
import uuid
import base64
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Any

logger = logging.getLogger(__name__)


class ConfigStorage:
    """Manages persistent storage of connections and templates."""
    
    def __init__(self):
        """Initialize config storage with default paths."""
        # Windows: Use AppData/Local, others: use home directory
        if os.name == 'nt':
            base_path = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        else:
            base_path = os.path.expanduser('~')
        
        self.config_dir = Path(base_path) / '.zabbix_extractor'
        self.config_file = self.config_dir / 'config.json'
        
        # Ensure directory exists
        self.config_dir.mkdir(parents=True, exist_ok=True)
        
        # Load or initialize config
        self.config = self._load_config()
    
    def _load_config(self) -> Dict:
        """Load config from file or return default structure."""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                logger.error(f"Error loading config: {e}")
                return self._default_config()
        return self._default_config()
    
    def _default_config(self) -> Dict:
        """Return default config structure."""
        return {
            'connections': [],
            'item_templates': []
        }
    
    def _save_config(self):
        """Save config to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except IOError as e:
            logger.error(f"Error saving config: {e}")
            raise
    
    # ============ Password Encoding (Basic obfuscation) ============
    
    def _encode_password(self, password: str) -> str:
        """Encode password with basic obfuscation (NOT secure encryption)."""
        if not password:
            return ""
        # Simple base64 encoding with prefix
        encoded = base64.b64encode(password.encode('utf-8')).decode('utf-8')
        return f"zb64:{encoded}"
    
    def _decode_password(self, encoded: str) -> str:
        """Decode password from obfuscated format."""
        if not encoded:
            return ""
        if encoded.startswith("zb64:"):
            try:
                return base64.b64decode(encoded[5:]).decode('utf-8')
            except Exception:
                return ""
        return encoded  # Return as-is if not encoded
    
    # ============ Connections Management ============
    
    def get_connections(self) -> List[Dict]:
        """Get all saved connections (with decoded passwords)."""
        connections = []
        for conn in self.config.get('connections', []):
            conn_copy = conn.copy()
            conn_copy['password'] = self._decode_password(conn.get('password', ''))
            connections.append(conn_copy)
        return connections
    
    def add_connection(self, name: str, url: str, username: str, password: str) -> Dict:
        """
        Add a new connection.
        
        Args:
            name: Display name for the connection
            url: Zabbix URL
            username: Login username
            password: Login password
            
        Returns:
            The created connection dict
            
        Raises:
            ValueError: If connection with same URL already exists
        """
        # Check for duplicate URL
        for conn in self.config['connections']:
            if conn['url'].lower() == url.lower():
                raise ValueError(f"Ya existe una conexión con URL: {url}")
        
        connection = {
            'id': str(uuid.uuid4()),
            'name': name,
            'url': url,
            'username': username,
            'password': self._encode_password(password),
            'created_at': datetime.now().strftime('%Y-%m-%d %H:%M')
        }
        
        self.config['connections'].append(connection)
        self._save_config()
        
        # Return copy with decoded password
        conn_copy = connection.copy()
        conn_copy['password'] = password
        return conn_copy
    
    def delete_connection(self, connection_id: str) -> bool:
        """Delete a connection by ID."""
        initial_count = len(self.config['connections'])
        self.config['connections'] = [
            c for c in self.config['connections'] if c['id'] != connection_id
        ]
        
        if len(self.config['connections']) < initial_count:
            self._save_config()
            return True
        return False
    
    def update_connection(self, connection_id: str, **kwargs) -> bool:
        """Update a connection's fields."""
        for conn in self.config['connections']:
            if conn['id'] == connection_id:
                for key, value in kwargs.items():
                    if key == 'password':
                        conn[key] = self._encode_password(value)
                    else:
                        conn[key] = value
                self._save_config()
                return True
        return False
    
    # ============ Item Templates Management ============
    
    def get_templates(self) -> List[Dict]:
        """Get all saved item templates."""
        return self.config.get('item_templates', [])
    
    def add_template(self, name: str, items: List[str], host_url: str = "") -> Dict:
        """
        Add a new item template linked to a host.
        
        Args:
            name: Template display name
            items: List of item names
            host_url: URL of the Zabbix host this template is linked to
            
        Returns:
            The created template dict
            
        Raises:
            ValueError: If template with same name exists or items empty
        """
        if not items:
            raise ValueError("Debe seleccionar al menos un item")
        
        if not name.strip():
            raise ValueError("El nombre del template no puede estar vacío")
        
        # Check for duplicate name within same host
        for tmpl in self.config['item_templates']:
            if tmpl['name'].lower() == name.lower() and tmpl.get('host_url', '').lower() == host_url.lower():
                raise ValueError(f"Ya existe un template con nombre: {name}")
        
        template = {
            'id': str(uuid.uuid4()),
            'name': name.strip(),
            'items': items,
            'host_url': host_url,
            'created_at': datetime.now().strftime('%Y-%m-%d %H:%M')
        }
        
        self.config['item_templates'].append(template)
        self._save_config()
        return template
    
    def delete_template(self, template_id: str) -> bool:
        """Delete a template by ID."""
        initial_count = len(self.config['item_templates'])
        self.config['item_templates'] = [
            t for t in self.config['item_templates'] if t['id'] != template_id
        ]
        
        if len(self.config['item_templates']) < initial_count:
            self._save_config()
            return True
        return False
    
    def get_template_by_id(self, template_id: str) -> Optional[Dict]:
        """Get a specific template by ID."""
        for tmpl in self.config['item_templates']:
            if tmpl['id'] == template_id:
                return tmpl
        return None
