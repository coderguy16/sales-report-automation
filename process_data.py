import pandas as pd
from openpyxl.chart import BarChart, LineChart, Reference
import os, re


def clean_sales_data(input_file='raw_sales_data.csv'):
    # Updated date parsing
    df = pd.read_csv(
        input_file,
        parse_dates=['order_date'],
        dayfirst=True,
        on_bad_lines='warn'
    )
    
    # Clean steps
    df = df.drop_duplicates().dropna(subset=['order_date'])
    
    
    
    # Rename quantity column
    df = df.rename(columns={
        'order_id': 'Order ID',
        'quantity': 'Quantity',
        'price': 'Price',
        'product': 'Product',
        'customer': 'Customer',
        'address': 'Address',
        'order_date': 'Order Date',
        'email': 'Email'
        
    })
    
    
    df['Order DateTime'] = pd.to_datetime(df['Order Date'], format='mixed')
    df['Order Date'] = df['Order DateTime'].dt.strftime('%Y-%m-%d')
    
    
    # Handle missing values
    numeric_cols = ['Quantity', 'Price']
    df[numeric_cols] = df[numeric_cols].fillna(0)
    df['Customer'] = df['Customer'].fillna('Unknown').str.strip()
    df['Email'] = df['Email'].fillna('Unknown')
    # Replace invalid with Invalid email given, so we know to contact the customer
    df['Email'] = df['Email'].replace('invalid', 'Invalid email given')
    # Transform Customer names to title case
    df['Customer'] = df['Customer'].str.title()
    df[['Street Name', 'City', 'State']] = df['Address'].apply(_parse_address)
    
    
    # Get everything after the last space in the address column as the zip code
    df['Zip Code'] = df['Address'].str.split().str[-1]
    
    # Clean product names
    product_mapping = {
        'laptop A': 'MacBook Pro',
        'laptop B': 'HP Envy',
        'laptop C': 'Dell Inspiron',
        'laptop D': 'Lenovo IdeaPad'
    }
    df['Product'] = df['Product'].replace(product_mapping).fillna('Other')
    
    # Add calculated column

    df['Total Sales'] = df['Quantity'] * df['Price']
    
    df['Formatted Total Sales'] = df['Total Sales'].apply(lambda x: "${:,.2f}".format(x))
    # Reorder the columns
    df = df [[
        'Order ID',
        'Customer',
        'Product',
        'Quantity',
        'Price',
        'Order Date',
        'Order DateTime',
        'Email',
        'Address',
        'Street Name',
        'City',
        'State',
        'Zip Code',
        'Total Sales',
        'Formatted Total Sales',
    ]]
    
    
    
    return df

def create_report(df, output_file=os.path.join(os.getcwd(), 'sales_report.xlsx')):
    # Create summary tables
    summary = df.pivot_table(
        index='Product',
        values=['Total Sales', 'Quantity'],
        aggfunc='sum'
    ).reset_index()
    
    # Updated frequency syntax to 'ME' (Month End)
    monthly_sales = df.set_index('Order DateTime').groupby(
        pd.Grouper(freq='ME')
    )['Total Sales'].sum()
    
    monthly_sales.index = monthly_sales.index.strftime('%B %Y')
    monthly_sales_df = monthly_sales.to_frame()
    
    # Format currency
    tables = [df, summary, monthly_sales_df]
    for d in tables:
        # check if 'price' or 'Total Sales' columns exist
        # Check if d is a dataframe:
        if   'price' in d.columns:
            d['Formatted Price'] = d['Price'].round(2).apply(lambda x: "${:,.2f}".format(x))
        d['Formatted Total Sales'] = d['Total Sales'].apply(lambda x: "${:,.2f}".format(x))
    
    # Create Excel report
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Cleaned Data', index=False)
        summary.to_excel(writer, sheet_name='Summary', index=False)
        monthly_sales_df.to_excel(writer, sheet_name='Monthly Sales')
        
        monthly_sales_df['date_datetime'] = pd.to_datetime(monthly_sales_df.index, format='%B %Y')

        
        # Add charts
        workbook = writer.book
        sheet = writer.sheets['Summary']
        
        # Create a bar chart
        chart1 = BarChart()
        chart1.type = "col"
        chart1.overlap = -10
        chart1.style = 10
        chart1.y_axis.number_format = '$#,##0'
        chart1.title = "Laptop Sales"
        chart1.y_axis.title = "Sales"
        chart1.x_axis.title = "Laptop"
        chart1.legend = None
        
        data = Reference(workbook['Summary'], min_col=3, min_row=1, max_col=3, max_row=len(summary)+1)
        
        cats = Reference(workbook['Summary'], min_col=1, min_row=2, max_row=len(summary)+1)
        chart1.add_data(data, titles_from_data=True)
        chart1.set_categories(cats)

        # Add the chart to the sheet
        sheet.add_chart(chart1, 'F2')
        
        
        # Line chart for monthly sales
        chart2 = LineChart()
        chart2.title = 'Monthly Sales'
        chart2.x_axis.title = 'Month'
        chart2.y_axis.title = 'Sales'
        chart2.y_axis.number_format = '$#,##0'
        chart2.style = 10
        chart2.overlap = -10
        chart2.legend = None
        
        data = Reference(workbook['Monthly Sales'], min_col=2, min_row=2, max_col=2, max_row=len(monthly_sales)+1)
        chart2.add_data(data)
    
        chart2.x_axis.number_format = 'mmmm yyyy'
        dates = Reference(workbook['Monthly Sales'], min_col=1, min_row=2, max_row=len(monthly_sales)+1)
        chart2.set_categories(dates)
        sheet.add_chart(chart2, 'F18')
        
            
        
    
    return output_file



# Function to extract street, city, and state
def _parse_address(address):
    # Updated Military Address Pattern
    military_pattern = re.compile(r"(PSC \d+(?:, Box \d+)?|Unit \d+(?:, Box \d+)?),?\s*(APO|FPO|DPO)\s+([A-Z]{2})\s+\d{5}")
    # military_pattern = re.compile(r"(PSC \d+|Box \d+),?\s*(APO|FPO|DPO)\s+([A-Z]{2})\s+\d{5}")
    standard_pattern = re.compile(r"([\d\w\s\.]+),\s*([\w\s]+),\s*([A-Z]{2})\s+\d{5}")

    # Check for Military Address
    military_match = military_pattern.search(address)
    if military_match:
        street_name = military_match.group(1)
        city = military_match.group(2)
        state = military_match.group(3)
        return pd.Series([street_name, city, state])

    # Check for Standard Address
    standard_match = standard_pattern.search(address)
    if standard_match:
        street_name = standard_match.group(1)
        city = standard_match.group(2)
        state = standard_match.group(3)
        return pd.Series([street_name, city, state])

    # Default case (if no match)
    return pd.Series([None, None, None])
