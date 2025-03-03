from faker import Faker
import pandas as pd
import random
from datetime import datetime, timedelta

fake = Faker()

def generate_messy_data(num_records=1000):
    products = ['laptop A', 'laptop B', 'laptop C', 'laptop D']
    
    data = []
    for _ in range(num_records):
        record = {
            # Intentionally messy data
            'Order ID': random.choice([fake.uuid4()]),
            'Customer': fake.name().upper(),  # All caps for some messiness
            'Product': random.choice(products),
            'Quantity': random.randint(1, 50),
            'Price': round(random.uniform(10, 500), 2),
            'Order Date': fake.date_between(start_date='-1y', end_date='today').strftime(
                random.choice(['%Y-%m-%d', '%m/%d/%Y', '%d-%b-%Y'])
            ),
            'Email': random.choice([fake.email(), 'invalid', '']),
            'Address': fake.address().replace('\n', ', ')
        }
        data.append(record)
    
    df = pd.DataFrame(data)
    
    # Introduce some null values and duplicates
    df.iloc[random.sample(range(num_records), 50)] = None
    df = pd.concat([df, df.sample(50)], ignore_index=True)
    
    # Save to CSV with inconsistent headers
    df.to_csv('raw_sales_data.csv', index=False, 
            header=['order_id', 'customer', 'product', 'quantity', 
                    'price', 'order_date', 'email', 'address'])
    
if __name__ == '__main__':
    generate_messy_data()
