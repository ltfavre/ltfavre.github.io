import sys
import pandas as pd
import sklearn
import pydantic

print("=" * 50)
print("KTEA/CELF LLM SCORING LAB")
print("=" * 50)

print(f"Python: {sys.version}")
print(f"Pandas: {pd.__version__}")
print(f"Scikit-Learn: {sklearn.__version__}")
print(f"Pydantic: {pydantic.__version__}")

print("\nEnvironment looks good.")