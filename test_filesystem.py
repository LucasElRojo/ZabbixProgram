"""
Test script para verificar get_filesystem_stats
"""

from pyzabbix import ZabbixAPI
import sys

# Connection settings
ZABBIX_URL = "http://192.172.1.159/zabbix/api_jsonrpc.php"
ZABBIX_USER = "Admin"
ZABBIX_PASS = "zabbix"

def main():
    print("=" * 60)
    print("TEST: get_filesystem_stats")
    print("=" * 60)
    
    # Connect
    print(f"\n1. Conectando...")
    api = ZabbixAPI(ZABBIX_URL)
    api.login(ZABBIX_USER, ZABBIX_PASS)
    print("   OK")
    
    # Get first host
    hosts = api.host.get(output=['hostid', 'name'], limit=1)
    if not hosts:
        print("ERROR: No hay hosts!")
        sys.exit(1)
    
    host = hosts[0]
    print(f"\n2. Host: {host['name']} (ID: {host['hostid']})")
    
    # Search for vfs.fs.size items
    print("\n3. Buscando items vfs.fs.size...")
    items = api.item.get(
        output=['itemid', 'name', 'key_', 'lastvalue', 'units'],
        hostids=host['hostid'],
        search={'key_': 'vfs.fs.size'},
        searchWildcardsEnabled=True,
        filter={'status': 0}
    )
    print(f"   Items vfs.fs.size: {len(items)}")
    
    if len(items) == 0:
        print("\n   ERROR: No se encontraron items vfs.fs.size!")
        print("   Probando busqueda alternativa por nombre...")
        items = api.item.get(
            output=['itemid', 'name', 'key_', 'lastvalue', 'units'],
            hostids=host['hostid'],
            search={'name': 'space'},
            filter={'status': 0}
        )
        print(f"   Items con 'space' en nombre: {len(items)}")
        
        # Show first 5 items
        print("\n   Primeros items encontrados:")
        for item in items[:5]:
            print(f"     - name: {item['name']}")
            print(f"       key_: {item['key_']}")
    
    # Parse filesystem data
    print("\n4. Parseando...")
    fs_data = {}
    parsed_count = 0
    
    for item in items:
        key = item.get('key_', '')
        lastvalue = item.get('lastvalue', '0')
        
        if 'vfs.fs.size[' not in key:
            continue
        
        try:
            params = key.split('[')[1].rstrip(']')
            parts = params.split(',')
            if len(parts) < 2:
                continue
            
            fsname = parts[0].strip()
            mode = parts[1].strip()
            
            parsed_count += 1
            
            if fsname not in fs_data:
                fs_data[fsname] = {'fsname': fsname}
            
            try:
                value = float(lastvalue)
            except:
                value = 0
            
            if mode == 'pused':
                fs_data[fsname]['pused'] = round(value, 1)
            elif mode == 'used':
                fs_data[fsname]['used_gb'] = round(value / (1024**3), 2)
            elif mode == 'total':
                fs_data[fsname]['total_gb'] = round(value / (1024**3), 2)
        except Exception as e:
            print(f"   Error: {e}")
    
    print(f"   Items parseados: {parsed_count}")
    
    # Filter results
    result = []
    for fsname, data in fs_data.items():
        if 'pused' in data:
            if 'used_gb' not in data and 'total_gb' in data:
                data['used_gb'] = round(data['total_gb'] * data['pused'] / 100, 2)
            result.append(data)
    
    result.sort(key=lambda x: x.get('fsname', ''))
    
    # Final validation
    print(f"\n5. FILESYSTEMS CON PUSED: {len(result)}")
    print("-" * 60)
    
    if len(result) == 0:
        print("   ERROR: 0 filesystems encontrados!")
        print("   La busqueda NO esta funcionando correctamente.")
        sys.exit(1)
    else:
        for fs in result:
            pused = fs.get('pused', 0)
            used_gb = fs.get('used_gb', 0)
            total_gb = fs.get('total_gb', 0)
            status = 'CRIT' if pused >= 90 else 'HIGH' if pused >= 80 else 'MED' if pused >= 70 else 'OK'
            print(f"   [{status:4}] {fs['fsname']:20} {pused:5.1f}% | {used_gb:7.2f}/{total_gb:7.2f} GB")
        
        print(f"\n   SUCCESS: {len(result)} filesystems encontrados!")

if __name__ == "__main__":
    main()
