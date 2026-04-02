import sys
import traceback

try:
    from invoice_app import main
    print("Starting application...")
    main()
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()
    input("Press Enter to exit...")
