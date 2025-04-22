import pdfplumber
import pandas as pd

data = []

with pdfplumber.open("Angebot_36227159-001.pdf") as pdf:
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split("\n"):
            if "POS" in line and "EUR" in line:
                data.append(line)

df = pd.DataFrame(data, columns=["Raw"])
df.to_csv("angebot_output.csv", index=False)
print("✅ Extraction complete.")
