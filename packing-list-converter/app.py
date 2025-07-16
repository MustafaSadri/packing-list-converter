import streamlit as st
import pandas as pd
from io import BytesIO

def process_packing_list(file):
    # Read Excel file
    df = pd.read_excel(file)

    # Extract SHIP TO details
    try:
        ship_to_info = df.iloc[6, 3]
    except:
        ship_to_info = "SHIP TO info not found"

    # Extract product data starting from row 14
    product_data = df.iloc[14:].copy()
    product_data = product_data.rename(columns={
        df.columns[1]: "Description",
        df.columns[-1]: "Quantity"
    })
    product_data = product_data[["Description", "Quantity"]].dropna(subset=["Description"])

    # Sales order details
    sales_order_no = df.iloc[2, -1] if not pd.isna(df.iloc[2, -1]) else "Unknown"
    sales_order_date = df.iloc[1, -1] if not pd.isna(df.iloc[1, -1]) else "Unknown"

    # Build final output
    output_df = pd.DataFrame()
    output_df["description"] = product_data["Description"]
    output_df["quantity"] = product_data["Quantity"]
    output_df["price"] = 1
    output_df["customer name"] = "Profit Development LLC"
    output_df["sales order no."] = sales_order_no
    output_df["sales order date"] = sales_order_date
    output_df["delivery method"] = "send by us"
    output_df["notes"] = ship_to_info

    return output_df

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Streamlit App UI
st.title("📦 Packing List → Sales Order Converter")

uploaded_file = st.file_uploader("Upload your packing list Excel (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success("✅ File uploaded successfully!")

    # Process the file
    sales_order_df = process_packing_list(uploaded_file)

    st.subheader("Preview of Generated Sales Order")
    st.dataframe(sales_order_df)

    # Download button
    excel_data = to_excel(sales_order_df)
    st.download_button(
        label="⬇️ Download Sales Order Excel",
        data=excel_data,
        file_name="Sales_Order_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )