import os
import sys
import glob
from datetime import datetime
from dateutil.relativedelta import relativedelta
from docx import Document

# 1. Legge i parametri passati da GitHub Actions
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

# 2. Calcolo automatico delle date degli instalment (+1, +2, +3 mesi)
instalment_1_date = subscription_creation_date
instalment_2_date = "N/A"
instalment_3_date = "N/A"
instalment_4_date = "N/A"

# Tenta il parsing della data nei formati più comuni (GG.MM.AAAA o GG/MM/AAAA)
for date_format in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
    try:
        base_date = datetime.strptime(subscription_creation_date, date_format)
        instalment_1_date = base_date.strftime("%d.%m.%Y")
        instalment_2_date = (base_date + relativedelta(months=1)).strftime("%d.%m.%Y")
        instalment_3_date = (base_date + relativedelta(months=2)).strftime("%d.%m.%Y")
        instalment_4_date = (base_date + relativedelta(months=3)).strftime("%d.%m.%Y")
        break
    except ValueError:
        continue

# 3. Ricerca del file template .docx
docx_files = glob.glob("*.docx")
if not docx_files:
    raise FileNotFoundError("Nessun file .docx trovato!")

template_path = docx_files[0]
doc = Document(template_path)

# 4. Mappatura completa di tutti i segnaposto
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
    # Segnaposto dinamici della Timeline
    "{INSTALMENT_1_DATE}": instalment_1_date,
    "{INSTALMENT_2_DATE}": instalment_2_date,
    "{INSTALMENT_3_DATE}": instalment_3_date,
    "{INSTALMENT_4_DATE}": instalment_4_date,
}

# 5. Sostituzione nei paragrafi
for p in doc.paragraphs:
    for key, value in replacements.items():
        if key in p.text:
            p.text = p.text.replace(key, value)

# 6. Sostituzione nelle tabelle
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for key, value in replacements.items():
                    if key in p.text:
                        p.text = p.text.replace(key, value)

# 7. Salva il documento aggiornato
output_filename = f"Rebuttal_{subscription_id}_{customer_name.replace(' ', '_')}.docx"
doc.save(output_filename)
print(f"File generato con successo: {output_filename}")
