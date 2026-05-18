import pandas as pd
import json
import time
import os
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeocoderServiceError

CACHE_FILE = 'geocache.json'

# Carregar cache existente
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, 'r', encoding='utf-8') as f:
        cache = json.load(f)
    print(f"Cache carregado: {len(cache)} endereços já geocodificados")
else:
    cache = {}

df = pd.read_excel('secretario.xlsx', sheet_name='DADOS', engine='openpyxl')
enderecos = df['MAPA'].dropna().unique().tolist()
print(f"Total de endereços únicos: {len(enderecos)}")

geolocator = Nominatim(user_agent="secretario_bc_dashboard_2026", timeout=10)

novos = 0
falhas = []

for i, addr in enumerate(enderecos):
    addr = str(addr).strip()
    if addr in cache:
        continue

    try:
        loc = geolocator.geocode(addr, exactly_one=True)
        if loc:
            cache[addr] = {'lat': loc.latitude, 'lon': loc.longitude}
            novos += 1
            print(f"[{i+1}/{len(enderecos)}] OK: {addr[:60]}")
        else:
            # Tentar só com rua + cidade
            partes = addr.split(',')
            addr_simples = partes[0].strip() + ', Balneário Camboriú, SC, Brasil'
            loc2 = geolocator.geocode(addr_simples, exactly_one=True)
            if loc2:
                cache[addr] = {'lat': loc2.latitude, 'lon': loc2.longitude}
                novos += 1
                print(f"[{i+1}/{len(enderecos)}] OK (simplificado): {addr[:60]}")
            else:
                cache[addr] = None
                falhas.append(addr)
                print(f"[{i+1}/{len(enderecos)}] FALHOU: {addr[:60]}")
        time.sleep(1.1)
    except (GeocoderTimedOut, GeocoderServiceError) as e:
        print(f"[{i+1}] Erro: {e} — tentando novamente...")
        time.sleep(3)
        try:
            loc = geolocator.geocode(addr, exactly_one=True)
            if loc:
                cache[addr] = {'lat': loc.latitude, 'lon': loc.longitude}
                novos += 1
            else:
                cache[addr] = None
                falhas.append(addr)
        except Exception:
            cache[addr] = None
            falhas.append(addr)
        time.sleep(1.1)

# Salvar cache
with open(CACHE_FILE, 'w', encoding='utf-8') as f:
    json.dump(cache, f, ensure_ascii=False, indent=2)

total_ok    = sum(1 for v in cache.values() if v is not None)
total_falha = sum(1 for v in cache.values() if v is None)
print(f"\nConcluído! {novos} novos geocodificados.")
print(f"Total com coordenadas: {total_ok} | Sem coordenadas: {total_falha}")
if falhas:
    print("Endereços que falharam:", falhas[:5])
