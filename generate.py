import sys
from docx import Document

# 1. Legge tutti i parametri passati da GitHub Actions
customer_name = sys.argv[1]
subscription_id = sys.argv[2]
subscription_creation_date = sys.argv[3]
shipping_order_id = sys.argv[4]
tracking_id = sys.argv[5]
shipping_total_value = sys.argv[6]
disputed_amount = sys.argv[7]
delivery_date = sys.argv[8]
consecutive_shipment_number = sys.argv[9]
disputed_charge_date = sys.argv[10]
disputed_instalment = sys.argv[11]
chargeback_reason = sys.argv[12]

# 2. Apri il file Word di template
template_path = "Chargeback rebuttal evidence and cover sheet template.docx"
doc = Document(template_path)

# 3. Mappatura corretta con le variabili Python valide
replacements = {
    "{CUSTOMER FULL NAME}": customer_name,
    "{SUBSCRIPTION ID}": subscription_id,
    "{SUBSCRIPTION DATE}": subscription_creation_date,
    "{ORDER ID}": shipping_order_id,
    "{TRACKING ID}": tracking_id,
    "{Shipping Value}": shipping_total_value,
    "{DISPUTE AMOUNT + CURRENCY}": disputed_amount,
    "{DELIVERY DATE}": delivery_date,
    "{shipping number}": consecutive_shipment_number,
    "{Disputed Charge Date}": disputed_charge_date,
    "{#X of 4}": disputed_instalment,
    "{e.g. Fraudulent / Not Received / Unauthorised}": chargeback_reason,
}

# 4. Sostituzione nei paragrafi
for p in doc.paragraphs:
    for key, value in replacements.items():
        if key in p.text:
            p.text = p.text.replace(key, value)

# 5. Sostituzione nelle tabelle
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for key, value in replacements.items():
                    if key in p.text:
                        p.text = p.text.replace(key, value)

# 6. Salva il documento
output_filename = f"Rebuttal_{subscription_id}_{customer_name.replace(' ', '_')}.docx"
doc.save(output_filename)
print(f"File generato con successo: {output_filename}")
