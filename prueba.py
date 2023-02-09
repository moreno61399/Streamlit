import streamlit as st

st.title("FK_KM CHECK")

fkonzept = st.file_uploader("upload FK file", type={"xlsx","csv", "txt"})

wb=workbook()

if fkonzept is not None:
    wb =load_workbook(fkonzept, read_only=True)
    st.write(wb.sheetnames)
    st.title(wb)
    st.write(wb.active)
    
st.download_button(
   "Press to Download your REPORT",
   wb,
   "file.csv",
   "text/csv",
   key='download-csv'
)
