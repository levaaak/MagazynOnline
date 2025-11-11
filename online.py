import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Skaner Magazynowy Online", layout="wide")

st.title("Skaner Magazynowy Online")

if "scanned_items" not in st.session_state:   
    st.session_state.scanned_items = []
if "found_codes" not in st.session_state:     
    st.session_state.found_codes = set()
if "double_scanned" not in st.session_state:  
    st.session_state.double_scanned = []
if "serial_input" not in st.session_state:     
    st.session_state.serial_input = ""
if "info_msg" not in st.session_state:
    st.session_state.info_msg = ""
if "highlight_color" not in st.session_state:
    st.session_state.highlight_color = "#F6F6F6"
if "pos_code" not in st.session_state:
    st.session_state.pos_code = ""
if "pos_name" not in st.session_state:
    st.session_state.pos_name = ""
if "inventory_date" not in st.session_state:
    st.session_state.inventory_date = datetime.now().strftime("%Y-%m-%d")

def handle_serial_change():
    kod = st.session_state.serial_input.strip()
    if not kod:
        return
    item_name = warehouse_items.get(kod, "")
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    found = kod in warehouse_items
    double_scan = found and (kod in st.session_state.found_codes)
    if found and not double_scan:
        status = "Zeskanowane"
        st.session_state.found_codes.add(kod)
        st.session_state.highlight_color = "#CBFFD4"
        st.session_state.info_msg = f"✅ Produkt {kod} ({item_name}) ZESKANOWANY"
    elif found and double_scan:
        status = "Podwójnie zeskanowane"
        st.session_state.highlight_color = "#FFA94D"
        st.session_state.info_msg = f"ℹ️ Produkt {kod} ({item_name}) został PODWÓJNIE zeskanowany!"
        st.session_state.double_scanned.append({"Kod produktu": kod,
                                                "Nazwa produktu": item_name,
                                                "Status": status,
                                                "Czas skanowania": now_str})
    else:
        status = "Dodatkowo zeskanowane"
        st.session_state.highlight_color = "#FFD2D2"
        st.session_state.info_msg = f"❌ Produkt {kod} nie jest w magazynie!"
    st.session_state.scanned_items.append({
        "Kod produktu": kod,
        "Nazwa produktu": item_name,
        "Status": status,
        "Czas skanowania": now_str
    })
    st.session_state.serial_input = ""

with st.sidebar:
    st.header("Parametry sesji")
    st.session_state.pos_code = st.text_input("Kod POS", value=st.session_state.pos_code)
    st.session_state.pos_name = st.text_input("Nazwa POS", value=st.session_state.pos_name)
    st.session_state.inventory_date = st.text_input("Data inwentaryzacji", value=st.session_state.inventory_date)

uploaded_file = st.file_uploader("Załaduj plik magazynowy (Excel z kolumnami: Kod produktu, Nazwa produktu)", type=['xlsx','xls'])
warehouse_items = {}
df_magazyn = None

if uploaded_file:
    try:
        df_magazyn = pd.read_excel(uploaded_file)
        col_code = df_magazyn.columns[0]
        col_name = df_magazyn.columns[1]
        for idx, row in df_magazyn.iterrows():
            code = str(row[col_code]).strip()
            name = str(row[col_name]).strip() if not pd.isna(row[col_name]) else ""
            if code:
                warehouse_items[code] = name
        st.success(f"Załadowano {len(warehouse_items)} pozycji magazynowych.")
        #st.dataframe(df_magazyn.head(10))
    except Exception as e:
        st.error(f"Problem z plikiem magazynu: {e}")

if warehouse_items:
    st.subheader("Skanuj lub wpisz kod produktu (ENTER dodaje do listy zeskanowanych):")
    st.text_input("Kod produktu", key="serial_input", on_change=handle_serial_change)
    st.markdown(
        f"<div style='background-color:{st.session_state.highlight_color};padding:20px;border-radius:8px;margin-bottom:10px;'>{st.session_state.info_msg}</div>",
        unsafe_allow_html=True
    )
    st.subheader("Historia skanów:")
    df_scanned = pd.DataFrame(st.session_state.scanned_items)
    st.dataframe(df_scanned, height=420)
    if st.button("Eksportuj raport do Excela"):
        scanned_codes = set(i['Kod produktu'] for i in st.session_state.scanned_items if i['Status'] in ["Zeskanowane", "Podwójnie zeskanowane"])
        not_found_in_scan = [code for code in warehouse_items if code not in scanned_codes]
        df_not_found = pd.DataFrame([
            {"Kod produktu": code, "Nazwa produktu": warehouse_items.get(code, "")}
            for code in not_found_in_scan
        ])
        df_extra = pd.DataFrame([
            i for i in st.session_state.scanned_items if i["Status"] == "Dodatkowo zeskanowane"
        ])
        df_double = pd.DataFrame(st.session_state.double_scanned)
        df_summary = pd.DataFrame({
            "W magazynie": [len(warehouse_items)],
            "Zeskanowane": [len(scanned_codes)],
            "Braki": [len(not_found_in_scan)],
            "Dodatkowe": [len(df_extra)]
        })
        pos_dict = {
            "Kod POS": [st.session_state.pos_code],
            "Nazwa POS": [st.session_state.pos_name],
            "Data inwentaryzacji": [st.session_state.inventory_date]
        }
        df_pos_info = pd.DataFrame(pos_dict)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_pos_info.to_excel(writer, sheet_name="Podsumowanie", index=False, startrow=0)
            df_summary.to_excel(writer, sheet_name="Podsumowanie", index=False, startrow=2)
            df_not_found.to_excel(writer, sheet_name="Niezeskanowane", index=False)
            df_scanned.to_excel(writer, sheet_name="Zeskanowane", index=False)
            df_extra.to_excel(writer, sheet_name="Dodatkowo zeskanowane", index=False)
            df_double.to_excel(writer, sheet_name="Podwójnie zeskanowane", index=False)
        st.download_button(
            label="Pobierz raport Excel",
            data=output.getvalue(),
            file_name=f"{st.session_state.pos_code}_{st.session_state.pos_name}_{st.session_state.inventory_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Załaduj plik magazynowy aby rozpocząć skanowanie.")