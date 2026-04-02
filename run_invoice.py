import sys
import traceback

try:
    print("Starting application...")
    from invoice_app import main
    main()
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()
    input("Press Enter to exit...")
