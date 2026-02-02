
import extract_msg
import sys
import os

print(f"extract-msg version: {getattr(extract_msg, '__version__', 'unknown')}")

try:
    # Use a dummy path or create a dummy msg if possible? 
    # Hard to create a dummy msg without a file. 
    # I will verify the class attributes directly.
    print("Attributes of extract_msg.Message:")
    print(dir(extract_msg.Message))
except Exception as e:
    print(e)
