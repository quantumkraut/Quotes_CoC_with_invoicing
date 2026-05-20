import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
from datetime import date, datetime, timedelta
from pathlib import Path
import os
from st_copy import copy_button

#======================= API
import requests


ACCESS_TOKEN = st.secrets["ACCESS_TOKEN"]
REFRESH_TOKEN = st.secrets["REFRESH_TOKEN"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]


def get_access_token():
    response = requests.post(
        "https://api.freeagent.com/v2/token_endpoint",
        data={
            "grant_type": "refresh_token",
            "refresh_token": REFRESH_TOKEN,
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
        }
    )

    tokens = response.json()
    return tokens["access_token"]


# ------------------ PAGE CONFIG ------------------

st.set_page_config(
    page_title="MTS Quote Generator",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# force white background feel (simple CSS)
st.markdown("""
    <style>

    /* MAIN APP BACKGROUND */
    .stApp {
        background-color: #0f172a;
        color: #e2e8f0;
    }

    /* GENERAL TEXT */
    p, span, div, label {
        color: #e2e8f0 !important;
    }

    /* HEADINGS */
    h1, h2, h3, h4 {
        color: #38bdf8 !important;
    }

    /* INPUT BOXES */
    input, textarea {
        background-color: #1e293b !important;
        color: #e2e8f0 !important;
        border: 1px solid #334155 !important;
    }

    /* SELECTBOX (IMPORTANT FIX FOR INVISIBLE DROPDOWN TEXT) */
    div[data-baseweb="select"] {
        background-color: #1e293b !important;
        color: #e2e8f0 !important;
    }

    div[data-baseweb="select"] * {
        color: #e2e8f0 !important;
    }

    /* DROPDOWN MENU */
    ul {
        background-color: #1e293b !important;
    }

    li {
        color: #e2e8f0 !important;
    }

    /* SIDEBAR */
    section[data-testid="stSidebar"] {
        background-color: #0b1220;
    }

    /* BUTTONS */
    /* MAIN STREAMLIT BUTTON */
    .stButton > button {
        background-color: #0f172a !important;  /* dark blue */
        color: #38bdf8 !important;             /* light blue text */
        border: 1px solid #38bdf8 !important;
    }

    /* HOVER STATE */
    .stButton > button:hover {
        background-color: #1e293b !important;
        color: #38bdf8 !important;
        border: 1px solid #38bdf8 !important;
    }

    /* ACTIVE / CLICKED STATE */
    .stButton > button:active {
        background-color: #0b1220 !important;
        color: #38bdf8 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------ INITIAL STATE ------------------

os.chdir(Path(__file__).parent)

if "custom_items" not in st.session_state:
    st.session_state.custom_items = {}

if "samples" not in st.session_state:
    st.session_state.samples = []

# ------------------ HEADER ------------------

st.title(":blue[MTS Quote & CoC Generator]")

st.divider()

# ------------------ LAB SELECTION ------------------

with st.container():
    st.subheader("Lab Selection")

    option = st.selectbox(
        "Which Lab is conducting the analysis?",
        ("DETS", "OEMA", "Envirochem", "Chemtech"),
    )


st.subheader("Client Selection")
client = st.text_input(label = "Who is the client?")

st.subheader("Project")
project = st.text_input(label = "What is the name of the project?")

st.subheader("Sampling Date")
date_of = st.date_input(label = "When was the sample collected?", format = "DD/MM/YYYY")

# ------------------ DETERMINAND SETUP ------------------

collectionfee = 0
determinands = {}

if option == "DETS":
    determinands = {
        "DETS Spectrum Suite": 133.51,
        "Cr(VI)": 4.82,
        "Total Sulphur": 6.28,
        "WAC (1-Stage)": 178.90,
        "HMRC LOI": 9.62,
        "Alkali Reserve": 24.94,
        "WS Chloride": 4.82,
        "WS Sulphates": 6.28,
        "Asbestos (Y/N)": 20.68,
        "Asbestos Quantification (gravimetric)": 45.76,
    }
    collectionfee = 40

elif option == "OEMA":
    determinands = {"WM3 Suite + Alkali Reserve Test + Asbestos ID": 337}
    collectionfee = 52.5

elif option == "Envirochem":
    determinands = {
        "SWTA Effluent Suite": 591.60,
        "Effluent Suite": 238.26,
        "MRF Dust Suite": 318.90,
        "Soil/Solid Suite": 329.76,
        "AD Grit WITH WAC": 420.24,
        "AD Grit": 193.14,
        "HMRC LOI": 6.28,
        "WAC (1-Stage)": 175.90,
        "Cr(VI)": 14.1,
    }
    collectionfee = 40

elif option == "Chemtech":
    determinands = {
        "Spectrum Suite": 125,
        "Full 1-Stage WAC": 150,
        "BS8601 Subsoils": 130.40,
        "BS3882 Topsoils": 135,
    }
    collectionfee = 26.25

# ------------------ DETERMINANDS UI ------------------

st.subheader("Number of Samples")
nosamp = st.selectbox(label = "How many samples?", options = ("1", "2", "3", "4", "5", "6", "7", "8", "9"))


st.subheader("Select Determinands")

cols = st.columns(3)
selected = {}

for i, (name, price) in enumerate(determinands.items()):
    with cols[i % 3]:
        if st.checkbox(f"{name} (£{price:.2f})"):
            selected[name] = price

# ------------------ CUSTOM ITEMS ------------------

st.subheader("Custom Determinands")

col1, col2 = st.columns(2)

with col1:
    custom_name = st.text_input("Name")

with col2:
    custom_price = st.number_input("Price (£)", min_value=1, step=1)

if st.button("Add to quote"):
    if custom_name:
        st.session_state.custom_items[custom_name] = custom_price

selected.update(st.session_state.custom_items)

# ------------------ TOTAL SUMMARY ------------------

total = sum(selected.values()) * int(nosamp)

st.subheader("Quote Summary")

c1, c2, c3 = st.columns(3)

with c1:
    if st.checkbox("Add collection fee", value=True):
        total += collectionfee

with c2:
    if st.checkbox("Add disposal fee (£2)"):
        total += 2 * int(nosamp)

with c3:
    st.metric("Total", f"£{total:.2f}")

st.write("---")





st.write(f"MTS Price (Total +25%): £{1.25 * total:.2f}")
st.header(f"MTS Price (incl. VAT): £{1.2 * 1.25 * total:.2f}")

#now to prepare to copy to clipboard
determinands_comma = "; ".join(selected)
#number calcs:
base = total
extra_25 = 0.25 * (base)
subtotal = base  + extra_25
extra_20 = 0.2 * subtotal
final_total = subtotal + extra_20

#to copy to clipbaord
tocopy = "\t".join([
    option,
    client,
    project,
    "",
    f"{nosamp} samples",
    determinands_comma,
    date_of.strftime("%d/%m/%Y"),
    str(base),
    str(collectionfee),
    str(extra_25),
    str(subtotal),
    str(extra_20),
    str(final_total),
    "No"
])


st.subheader(":red[Click below to copy for MTS Systems]")
copy_button(tocopy, tooltip="Copy", copied_label="Copied Successfully!", icon="st")

st.write("---")
#------------------ Invoicing ------------------


######### FREE AGENT API INVOICE SECTION

st.header(":red[Invoice Generator]")

import requests

today = date.today()
duedate = today + timedelta(days=30)

todaystr = today.strftime("%Y-%m-%d")
duedatestr = duedate.strftime("%Y-%m-%d")


## Get all contacts into dictionary
#function
def get_contacts():

    access_token = get_access_token()

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    r = requests.get(
        "https://api.freeagent.com/v2/contacts",
        headers=headers
    )

    return r.json()

#calling fucntion and storing values
contacts_data = get_contacts()

contacts_dict = {}

for contact in contacts_data["contacts"]:

    if "organisation_name" in contact:
        name = contact["organisation_name"]

    else:
        name = f"{contact['first_name']} {contact['last_name']}"

    contacts_dict[name] = contact["url"]

freeagent_contacts = []
for i in contacts_dict:
    freeagent_contacts.append(i)

contact_button = st.selectbox(label = "Select Contact for Invoice", options=freeagent_contacts, placeholder="Select from dropdown", index=None)


if contact_button != None:
    clienturl = contacts_dict[contact_button]
    if st.button("Create Invoice"):
        reference = datetime.now().strftime("%d%m%Y%H%M")

        access_token = get_access_token()

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        invoice = {
            "invoice": {
                "contact": clienturl,
                "dated_on": todaystr,
                "payment_terms_in_days": 30,
                "reference": reference,
                "invoice_items": [
                    {
                        "description": f"Lab testing: {determinands_comma}",
                        "quantity": 1,
                        "price": subtotal
                    }
                ]
            }
        }

        r = requests.post(
            "https://api.freeagent.com/v2/invoices",
            json=invoice,
            headers=headers
        )

        data = r.json()

        st.success("Invoice created successfully")



# ------------------ COC SECTION ------------------
st.write("---")
st.write("")

st.header(":red[CoC Generation]")

template_map = {
    "DETS": "DETs 2026 COC.xlsx",
    "Envirochem": "EnvirochemCoC.xlsx",
    "Chemtech": "ChemtechCoC.xlsx",
}

filename = template_map.get(option)

if not filename:
    st.warning("No template available")
    st.stop()

projectCOC = st.text_input("Project Name", value = project)
quotation = st.text_input("Quotation Number", value="Q25-20566 & Q26-10205")

# ------------------ SAMPLES ------------------

mapping = {
    "Soil": "S",
    "Miscellaneous": "M",
    "Trade Effluent": "EF",
    "Groundwater": "GW",
    "River/Surface Water": "RW",
}

if st.button("➕ Add Sample"):
    st.session_state.samples.append({"id": "", "type": "Soil", "date": None})

for i, sample in enumerate(st.session_state.samples):
    st.markdown(f"### Sample {i+1}")

    sample["id"] = st.text_input("Sample ID", key=f"id_{i}", value=f"Sample {i+1}")
    sample["type"] = st.selectbox("Type", list(mapping.keys()), key=f"type_{i}")
    sample["date"] = st.date_input("Date", key=f"date_{i}", format="DD/MM/YYYY")


st.write("")


# ------------------ GENERATE ------------------

if st.button("Generate CoC"):

    wb = load_workbook(filename)
    ws = wb.active

    ws["L3"] = f"Contract Title: {projectCOC}"
    ws["L5"] = f"Client Ref: {projectCOC}"
    ws["L7"] = f"Quotation No: {quotation}"

    cols = ["AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL"]
    sample_x_cols = ["AB", "AD"]
    selected_list = list(selected.keys())

    for i, det in enumerate(selected_list):
        if i < len(cols):
            ws[f"{cols[i]}4"] = det

    start_row = 15

    for s_idx, sample in enumerate(st.session_state.samples):
        row = start_row + s_idx

        ws[f"B{row}"] = sample["id"]
        ws[f"D{row}"] = sample["type"]
        ws[f"E{row}"] = mapping[sample["type"]]
        ws[f"I{row}"] = sample["date"]

        for c in sample_x_cols:
            ws[f"{c}{row}"] = "X"

        for i in range(len(selected_list)):
            if i < len(cols):
                ws[f"{cols[i]}{row}"] = "X"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    today = date.today().strftime("%Y-%m-%d")

    st.download_button(
        "Download CoC",
        buffer,
        file_name=f"{today}_{projectCOC}_{option}_CoC.xlsx".replace(" ", "_"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



