try:
    from vnstock import *
    print("vnstock imported successfully")
except Exception as e:
    print("Error importing vnstock:", e)
    
try:
    # the function is usually exchange_rate
    rates = exchange_rate('USD', '2025-01-02', '2025-01-05')
    print("USD rates:")
    print(rates)
except Exception as e:
    print("Error getting historical rates from vnstock:", e)
