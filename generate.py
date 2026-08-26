import sys
from docx import Document

# Legge gli argomenti passati dal Form di GitHub
customer_name = sys.argv[1]
subscription_id = sys.argv[2]
disputed_amount = sys.argv[3]
delivery_date = sys.argv[4]

template_path = "template.docx"
doc = Document(template_path)

replacements = {
    "{CUSTOMER FULL NAME}": customer_name,
    "{SUBSCRIPTION ID}": subscription_id,
    "{SUBSCRIPTION DATE}" : subscription_creation_date,
    "{ORDER ID}": Shipping Order ID,
    "{TRACKING ID}": Tracking ID,
    "{Shipping Value}": Shipping Total Value,
    "{DISPUTE AMOUNT AMOUNT + CURRENCY}": disputed_amount,
    "{DELIVERY DATE}": delivery_date,
    "{shipping number}": Consecutive Shipment Number,
    "{Disputed Charge Date}": Disputed Charge Date,
    "{#X of 4}": Disputed Instalment,
    "{e.g. Fraudulent / Not Received / Unauthorised}": Chargeback Reason,
}

# Sostituisce i dati nei paragrafi
for p in doc.paragraphs:
    for key, value in replacements.items():
        if key in p.text:
            p.text = p.text.replace(key, value)

# Sostituisce i dati nelle tabelle
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for key, value in replacements.items():
                    if key in p.text:
                        p.text = p.text.replace(key, value)

output_filename = (
    f"Rebuttal_{subscription_id}_{customer_name.replace(' ', '_')}.docx"
)
doc.save(output_filename)
print(f"File generato: {output_filename}")
