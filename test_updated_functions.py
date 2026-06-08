import sys
sys.path.insert(0, r'D:\Tygia-Tudong\scripts')
from get_rates import fetch_vietinbank_usd, fetch_seabank_usd, fetch_techcombank_usd

print("Testing VietinBank Server Action API...")
vtb = fetch_vietinbank_usd()
print(f"VietinBank USD: {vtb}")

print("\nTesting SeaBank Server Action API...")
seab = fetch_seabank_usd()
print(f"SeaBank USD: {seab}")

print("\nTesting Techcombank session API...")
tcb = fetch_techcombank_usd()
print(f"Techcombank USD: {tcb}")
