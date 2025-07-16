import streamlit as st
import pandas as pd
from io import BytesIO

def process_packing_list(file):
    try:
        # Determine engine based on file extension
        engine = 'openpyxl' if file.name.endswith('.xlsx') else 'xlrd'
        
        # Read Excel file with explicit engine
        df = pd.read_excel(file, engine=engine)

        # Extract SHIP TO details with better error handling
        try:
            ship_to_info = str(df.iloc[6, 3])  # Convert to string in case it's not
        except Exception as e:
            st.warning(f"Could not extract SHIP TO info: {str(e)}")
            ship_to_info = "SHIP TO info not found"

        # Extract product data with more robust column handling
        product_data = df.iloc[14:].copy()
        
        # Safely rename columns
        description_col = df.columns[1] if len(df.columns) > 1 else None
        quantity_col = df.columns[-1]
        
        if not description_col:
            raise ValueError("Excel file doesn't have expected columns")
            
        product_data = product_data.rename(columns={
            description_col: "Description",
            quantity_col: "Quantity"
        })
        
        # Filter and clean data
        product_data = product_data[["Description", "Quantity"]].dropna(subset=["Description"])
        product_data = product_data[product_data["Description"].astype(str).str.strip() != ""]

        # Sales order details with better null handling
        sales_order_no = str(df.iloc[2, -1]) if len(df.columns) > 0 and not pd.isna(df.iloc[2, -1]) else "Unknown"
        sales_order_date = str(df.iloc[1, -1]) if len(df.columns) > 0 and not pd.isna(df.iloc[1, -1]) else "Unknown"

        # Build final output
        output_df = pd.DataFrame({
            "description": product_data["Description"].astype(str),
            "quantity": pd.to_numeric(product_data["Quantity"], errors='coerce').fillna(0),
            "price": 1,
            "customer name": "Profit Development LLC",
            "sales order no.": sales_order_no,
            "sales order date": sales_order_date,
            "delivery method": "send by us",
            "notes": ship_to_info
        })

        return output_df

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return pd.DataFrame()  # Return empty DataFrame on error

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
