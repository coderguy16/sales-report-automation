from process_data import clean_sales_data, create_report
from email_report import send_email
import generate_data  # Only needed for initial data generation

if __name__ == '__main__':
    # Step 1: Generate data (only needed once)
    generate_data.generate_messy_data()
    
    # Step 2: Process data
    cleaned_df = clean_sales_data()
    report_file = create_report(cleaned_df)
    
    print('Report generated under:', report_file)
    
    # Step 3: Send email
    try:
        send_email(report_file)
        print("Report generated and sent successfully!")
    except Exception as e:
        print(f"Email sending failed: {e}")
