import streamlit as st
import pandas as pd
import smartsheet
import re

APP_NAME = 'Air Filter Ordering'


@st.cache_data(ttl=300)
def smartsheet_to_dataframe(sheet_id):
    smartsheet_client = smartsheet.Smartsheet(st.secrets['smartsheet']['access_token'])
    sheet             = smartsheet_client.Sheets.get_sheet(sheet_id)
    columns           = [col.title for col in sheet.columns]
    rows              = []
    for row in sheet.rows: rows.append([cell.value for cell in row.cells])
    return pd.DataFrame(rows, columns=columns)

def expand_sizes(s):
    result = []
    matches = re.findall(r'(\d+):\(([^)]+)\)', s)
    for qty, size in matches:
        result.extend([size] * int(qty))
    return result

st.set_page_config(page_title=APP_NAME, page_icon='💨', layout='centered')

st.image(st.secrets['images']["rr_logo"], width=100)

st.title(APP_NAME)
st.info('Creation of an air filter order using the filters currently in homes and a warehouse count.')

df                          = smartsheet_to_dataframe(st.secrets['smartsheet']['sheets']['schedule'])
df                          = df.dropna(subset=['Filters'])
df['Expanded']              = df['Filters'].apply(expand_sizes)
filter_sums                 = sum(df['Expanded'], [])
df_counts                   = pd.Series(filter_sums).value_counts().reset_index()
df_counts.columns           = ['Size', 'Count']
df_counts                   = df_counts[df_counts["Size"].str.match(r'^\d+(?:\.\d+)? X \d+(?:\.\d+)? X \d+(?:\.\d+)?$')]
df_counts[["L","W","H"]]    = df_counts["Size"].str.split(" X ", expand=True)
df_counts                   = df_counts.sort_values(by=["L","W"]).reset_index(drop=True)
df_counts                   = df_counts[["Size","Count"]]

st.subheader('Currently In Homes')
st.dataframe(df_counts, hide_index=True, width='stretch')

df_needed = df_counts.copy()
df_needed['Count'] = None

st.download_button('Download Needed Filters Template', df_needed.to_csv(index=False).encode('utf-8'), file_name='needed_filters_template.csv', mime='text/csv', width='stretch')

st.subheader('Current Month Still To Be Changed')
weeks = st.multiselect('Weeks', options=df['Week'].astype(int).unique(), default=df['Week'].astype(int).unique().tolist())

df_weeks = df[df['Week'].astype(int).isin(weeks)]
week_filter_sums = sum(df_weeks['Expanded'], [])
df_week_counts = pd.Series(week_filter_sums).value_counts().reset_index()
df_week_counts.columns = ['Size', 'Count']
df_week_counts = df_week_counts[df_week_counts["Size"].str.match(r'^\d+(?:\.\d+)? X \d+(?:\.\d+)? X \d+(?:\.\d+)?$')]
df_week_counts[["L","W","H"]] = df_week_counts["Size"].str.split(" X ", expand=True)
df_week_counts = df_week_counts.sort_values(by=["L","W"]).reset_index(drop=True)
df_week_counts = df_week_counts[["Size","Count"]]

st.subheader('Currently In Warehouse')
uploaded_count = st.file_uploader('Current Warehouse Inventory', type='csv')

if uploaded_count is not None:

    st.subheader('Order Summary')
    df_uploaded = pd.read_csv(uploaded_count)
    
    df = pd.merge(df_counts, df_uploaded, on='Size', how='left', suffixes=('_In_Homes', '_In_Warehouse'))
    df = pd.merge(df, df_week_counts, on='Size', how='left')
    df['Count_In_Warehouse'] = df['Count_In_Warehouse'].fillna(0).astype(int)
    df['Count'] = df['Count'].fillna(0).astype(int)
    df['Remainder'] = df['Count_In_Warehouse'] - df['Count_In_Homes'].fillna(0) - df['Count'].fillna(0)
    df = df[['Size', 'Count_In_Warehouse', 'Count', 'Count_In_Homes', 'Remainder']]
    df.columns = ['Size', 'In Warehouse', 'Needed This Month', 'Needed Next Month', 'Remainder']
    
    
    df_order = df[df['Remainder'] < 0].copy()
    df_order.rename(columns={'Remainder': 'Filters_Needed'}, inplace=True)
    df_order['Filters_Needed'] = df_order['Filters_Needed'].abs()
    df_order = df_order[['Size', 'Filters_Needed']]

    filters  = smartsheet_to_dataframe(st.secrets['smartsheet']['sheets']['filters'])

    df_order = pd.merge(df_order, filters, on='Size', how='left')
    df_order['Cases_Needed'] = (df_order['Filters_Needed'] / df_order['Quantity_Per_Case']).apply(lambda x: int(x) + (1 if x % 1 > 0 else 0))

    with st.expander('Details'):
        '**Determining Remaining Inventory**'
        st.dataframe(df, hide_index=True, width='stretch')

        '**Determining Cases Needed**'
        st.dataframe(df_order, hide_index=True, width='stretch')

    for vendor, df_vendor in df_order.groupby('Vendor'):
        f'**{vendor}**'

        df_vendor = df_vendor[['Size','Cost','Cases_Needed']].copy()
        df_vendor.rename(columns={'Cases_Needed': 'Cases'}, inplace=True)
        df_vendor['Total_Cost'] = df_vendor['Cost'] * df_vendor['Cases']

        l, r = st.columns(2)
        l.metric('Total Cases', df_vendor['Cases'].sum())
        r.metric('Total Cost', f"${df_vendor['Total_Cost'].sum():,.2f}")
        
        st.dataframe(df_vendor, hide_index=True, width='stretch')        
        st.download_button(f'Download Order for **{vendor}**', df_vendor[['Size','Cases']].to_csv(index=False).encode('utf-8'), file_name=f'order_{vendor.lower().replace(' ','_')}.csv', mime='text/csv', width='stretch', type='primary', key=vendor+'_download')