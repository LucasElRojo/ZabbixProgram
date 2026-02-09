"""
Zabbix API Client Module
Handles authentication and data fetching from Zabbix server.
"""

from pyzabbix import ZabbixAPI
from typing import Optional, List, Dict, Any
import logging

logger = logging.getLogger(__name__)


class ZabbixClient:
    """Client for interacting with Zabbix API."""
    
    def __init__(self):
        self.api: Optional[ZabbixAPI] = None
        self.base_url: str = ""
        self.session_id: str = ""
        self._connected: bool = False
    
    @property
    def is_connected(self) -> bool:
        """Check if client is connected to Zabbix."""
        return self._connected
    
    def connect(self, url: str, user: str, password: str) -> bool:
        """
        Connect and authenticate with Zabbix server.
        
        Args:
            url: Zabbix server URL (e.g., https://zabbix.example.com)
            user: Username for authentication
            password: Password for authentication
            
        Returns:
            True if connection successful, False otherwise
            
        Raises:
            Exception: If connection or authentication fails
        """
        import requests
        from requests.exceptions import ConnectionError, Timeout, RequestException
        from pyzabbix import ZabbixAPIException
        
        try:
            # Normalize URL
            self.base_url = url.rstrip('/')
            if not self.base_url.endswith('/api_jsonrpc.php'):
                api_url = f"{self.base_url}/api_jsonrpc.php"
            else:
                api_url = self.base_url
                self.base_url = self.base_url.replace('/api_jsonrpc.php', '')
            
            logger.info(f"Connecting to Zabbix at {api_url}")
            
            self.api = ZabbixAPI(api_url)
            self.api.timeout = 10  # Set timeout to 10 seconds
            self.api.login(user, password)
            self.session_id = self.api.auth
            self._connected = True
            
            logger.info(f"Successfully connected. Session ID: {self.session_id[:8]}...")
            return True
            
        except ZabbixAPIException as e:
            self._connected = False
            error_str = str(e).lower()
            if 'incorrect' in error_str or 'invalid' in error_str or 'login' in error_str:
                error_msg = "Usuario o contraseña incorrectos"
            elif 'permission' in error_str:
                error_msg = "Sin permisos de acceso"
            else:
                error_msg = f"Error de API Zabbix: {str(e)}"
            logger.error(f"Zabbix API error: {str(e)}")
            raise Exception(error_msg)
            
        except ConnectionError as e:
            self._connected = False
            error_msg = f"No se pudo conectar al servidor. Verifique que la URL sea correcta y el servidor esté accesible."
            logger.error(f"Connection error: {str(e)}")
            raise Exception(error_msg)
            
        except Timeout as e:
            self._connected = False
            error_msg = "Tiempo de espera agotado. El servidor no responde."
            logger.error(f"Timeout error: {str(e)}")
            raise Exception(error_msg)
            
        except RequestException as e:
            self._connected = False
            error_msg = f"Error de red: {str(e)}"
            logger.error(f"Request error: {str(e)}")
            raise Exception(error_msg)
            
        except Exception as e:
            self._connected = False
            error_str = str(e).lower()
            if 'connection refused' in error_str:
                error_msg = "Conexión rechazada. Verifique que el servidor Zabbix esté ejecutándose."
            elif 'name or service not known' in error_str or 'nodename nor servname' in error_str:
                error_msg = "No se pudo resolver el nombre del servidor. Verifique la URL."
            elif 'ssl' in error_str or 'certificate' in error_str:
                error_msg = "Error de certificado SSL. Intente con http:// en lugar de https://"
            else:
                error_msg = f"Error de conexión: {str(e)}"
            logger.error(f"Connection failed: {str(e)}")
            raise Exception(error_msg)
    
    def disconnect(self) -> None:
        """Disconnect from Zabbix server."""
        if self.api and self._connected:
            try:
                self.api.user.logout()
            except Exception:
                pass
        self._connected = False
        self.api = None
        self.session_id = ""
        logger.info("Disconnected from Zabbix")
    
    def get_templates(self) -> List[Dict[str, Any]]:
        """
        Get all available templates.
        
        Returns:
            List of templates with templateid and name
        """
        if not self._connected or not self.api:
            raise Exception("No conectado al servidor Zabbix")
        
        try:
            templates = self.api.template.get(
                output=['templateid', 'name'],
                sortfield='name'
            )
            logger.info(f"Found {len(templates)} templates")
            return templates
        except Exception as e:
            logger.error(f"Error fetching templates: {str(e)}")
            raise Exception(f"Error al obtener templates: {str(e)}")
    
    def get_hosts_by_template(self, template_id: str) -> List[Dict[str, Any]]:
        """
        Get hosts linked to a specific template.
        
        Args:
            template_id: The template ID to filter by
            
        Returns:
            List of hosts with hostid and name
        """
        if not self._connected or not self.api:
            raise Exception("No conectado al servidor Zabbix")
        
        try:
            hosts = self.api.host.get(
                output=['hostid', 'name'],
                templateids=template_id,
                sortfield='name'
            )
            logger.info(f"Found {len(hosts)} hosts for template {template_id}")
            return hosts
        except Exception as e:
            logger.error(f"Error fetching hosts: {str(e)}")
            raise Exception(f"Error al obtener hosts: {str(e)}")
    
    def get_items_by_host(self, host_id: str) -> List[Dict[str, Any]]:
        """
        Get monitored items for a specific host.
        
        Args:
            host_id: The host ID to get items from
            
        Returns:
            List of items with itemid, name, and key_
        """
        if not self._connected or not self.api:
            raise Exception("No conectado al servidor Zabbix")
        
        try:
            items = self.api.item.get(
                output=['itemid', 'name', 'key_', 'value_type'],
                hostids=host_id,
                sortfield='name',
                filter={'status': 0}  # Only enabled items
            )
            # Filter items that can generate graphs (numeric values)
            graphable_items = [
                item for item in items 
                if item.get('value_type') in ['0', '3']  # 0=float, 3=unsigned int
            ]
            logger.info(f"Found {len(graphable_items)} graphable items for host {host_id}")
            return graphable_items
        except Exception as e:
            logger.error(f"Error fetching items: {str(e)}")
            raise Exception(f"Error al obtener items: {str(e)}")
    
    def get_items_by_template(self, template_id: str) -> List[Dict[str, Any]]:
        """
        Get items defined in a template (not inherited).
        These items are common to all hosts linked to this template.
        
        Args:
            template_id: The template ID to get items from
            
        Returns:
            List of items with name and key_
        """
        if not self._connected or not self.api:
            raise Exception("No conectado al servidor Zabbix")
        
        try:
            items = self.api.item.get(
                output=['itemid', 'name', 'key_', 'value_type'],
                templateids=template_id,
                sortfield='name',
                filter={'status': 0}
            )
            # Filter items that can generate graphs (numeric values)
            graphable_items = [
                item for item in items 
                if item.get('value_type') in ['0', '3']  # 0=float, 3=unsigned int
            ]
            logger.info(f"Found {len(graphable_items)} graphable items for template {template_id}")
            return graphable_items
        except Exception as e:
            logger.error(f"Error fetching template items: {str(e)}")
            raise Exception(f"Error al obtener items del template: {str(e)}")
    
    def get_session_cookie(self) -> str:
        """
        Get the session cookie for chart requests.
        
        Returns:
            Session ID string for use in cookies
        """
        if not self._connected:
            raise Exception("No conectado al servidor Zabbix")
        return self.session_id
    
    def get_base_url(self) -> str:
        """
        Get the base URL of the Zabbix server.
        
        Returns:
            Base URL string
        """
        return self.base_url
    
    def get_filesystem_stats(self, host_id: str) -> List[Dict[str, Any]]:
        """
        Get filesystem/partition statistics for a host.
        
        Args:
            host_id: The Zabbix host ID
            
        Returns:
            List of dicts: [{fsname, pused, used_gb, total_gb, free_gb}, ...]
        """
        if not self._connected or not self.api:
            raise Exception("No conectado al servidor Zabbix")
        
        try:
            # Get filesystem items using key pattern vfs.fs.size
            items = self.api.item.get(
                output=['itemid', 'name', 'key_', 'lastvalue', 'units'],
                hostids=host_id,
                search={'key_': 'vfs.fs.size'},
                searchWildcardsEnabled=True,
                filter={'status': 0}
            )
            
            logger.info(f"Found {len(items)} vfs.fs.size items for host {host_id}")
            
            if not items:
                return []
            
            # Group items by filesystem name
            fs_data = {}  # fsname -> {pused, used, total}
            
            for item in items:
                key = item.get('key_', '')
                lastvalue = item.get('lastvalue', '0')
                
                # Parse key: vfs.fs.size[/path,mode]
                if 'vfs.fs.size[' not in key:
                    continue
                
                try:
                    # Extract parameters from key
                    params = key.split('[')[1].rstrip(']')
                    parts = params.split(',')
                    if len(parts) < 2:
                        continue
                        
                    fsname = parts[0].strip()
                    mode = parts[1].strip()
                    
                    if fsname not in fs_data:
                        fs_data[fsname] = {'fsname': fsname}
                    
                    # Convert value
                    try:
                        value = float(lastvalue)
                    except (ValueError, TypeError):
                        value = 0
                    
                    # Store by mode
                    if mode == 'pused':
                        fs_data[fsname]['pused'] = round(value, 1)
                    elif mode == 'used':
                        fs_data[fsname]['used_gb'] = round(value / (1024**3), 2)
                    elif mode == 'total':
                        fs_data[fsname]['total_gb'] = round(value / (1024**3), 2)
                    elif mode == 'free':
                        fs_data[fsname]['free_gb'] = round(value / (1024**3), 2)
                        
                except Exception as parse_err:
                    logger.debug(f"Could not parse key {key}: {parse_err}")
                    continue
            
            # Filter to only include filesystems with pused data
            result = []
            for fsname, data in fs_data.items():
                if 'pused' in data:
                    # Calculate missing values if possible
                    if 'used_gb' not in data and 'total_gb' in data and 'pused' in data:
                        data['used_gb'] = round(data['total_gb'] * data['pused'] / 100, 2)
                    if 'free_gb' not in data and 'total_gb' in data and 'used_gb' in data:
                        data['free_gb'] = round(data['total_gb'] - data['used_gb'], 2)
                    result.append(data)
            
            # Sort by filesystem name
            result.sort(key=lambda x: x.get('fsname', ''))
            logger.info(f"Returning {len(result)} filesystems with pused data")
            return result
            
        except Exception as e:
            logger.error(f"Error fetching filesystem stats: {str(e)}")
            return []
