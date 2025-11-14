import duckdb
import pandas as pd
from datetime import datetime
import os

def load_stock_master():
    """Load the Stock Master file to get order type information."""
    stock_file = 'Stock master Total.xlsx'
    if not os.path.exists(stock_file):
        raise FileNotFoundError(f"Stock Master file not found: {stock_file}")
    
    # Load the Stock Master file
    stock_df = pd.read_excel(stock_file, usecols=['STOCK', 'AC_GROUP'])
    
    # Clean and standardize the data
    stock_df = stock_df.dropna(subset=['STOCK', 'AC_GROUP'])
    stock_df['STOCK'] = stock_df['STOCK'].astype(str).str.strip()
    stock_df['AC_GROUP'] = stock_df['AC_GROUP'].astype(str).str.strip().str.upper()
    
    # Filter only valid order types
    valid_types = ['DINE-IN', 'TAKEAWAY', 'DELIVERY']
    stock_df = stock_df[stock_df['AC_GROUP'].isin(valid_types)]
    
    # Create a mapping of stock code to order type
    stock_to_order_type = dict(zip(stock_df['STOCK'], stock_df['AC_GROUP']))
    return stock_to_order_type

def calculate_store_sales(store_code='204'):
    """
    Calculate store sales by summing up item details.
    
    Args:
        store_code (str): The store code to calculate sales for
        
    Returns:
        pd.DataFrame: DataFrame containing the sales summary
    """
    db_path = ".pos_cache/sushi_epoint_pos_live/pos_20240102.duckdb"
    
    try:
        # Connect to the database
        conn = duckdb.connect(database=db_path, read_only=True)

        # Load stock master sales categories (Column I / AC_GROUP)
        stock_df = pd.read_excel(
            'Stock master Total.xlsx',
            usecols=['STOCK', 'AC_GROUP']
        ).dropna(subset=['STOCK', 'AC_GROUP'])
        stock_df = stock_df.rename(
            columns={'STOCK': 'stock_code', 'AC_GROUP': 'sales_category'}
        )
        stock_df['stock_code'] = stock_df['stock_code'].astype(str).str.strip()
        stock_df['sales_category'] = (
            stock_df['sales_category'].astype(str).str.strip().str.upper()
        )
        conn.register('stock_master', stock_df)

        # Query sales data using stock master sales categories
        query = f"""
        WITH sales_data AS (
            SELECT
                d.sales_no AS receipt_no,
                d.item_name,
                d.qty,
                COALESCE(d.sub_total, 0) AS gross_amount,
                COALESCE(d.svc_amt, 0) AS service_charge,
                COALESCE(d.tax_amt, 0) AS item_tax,
                TRIM(CAST(d.item_no AS VARCHAR)) AS item_code,
                r.terminal,
                r.date AS receipt_date,
                r.time AS receipt_time,
                sm.sales_category,
                CASE
                    WHEN sm.sales_category IN ('DINE-IN', 'TAKEAWAY', 'DELIVERY')
                        THEN sm.sales_category
                    WHEN COALESCE(d.svc_amt, 0) > 0 THEN 'DINE-IN'
                    WHEN COALESCE(d.take_away_item, '') = 'Y' THEN 'TAKEAWAY'
                    WHEN r.terminal IN ('T004', 'T005', 'T006') THEN 'DINE-IN'
                    WHEN r.terminal = 'T001' THEN 'TAKEAWAY'
                    ELSE 'UNKNOWN'
                END AS final_order_type
            FROM raw_details d
            JOIN raw_receipts r ON d.sales_no = r.receipt_no
            LEFT JOIN stock_master sm
                ON TRIM(CAST(d.item_no AS VARCHAR)) = sm.stock_code
            WHERE d.store_code = '{store_code}'
              AND r.is_void = 0
              AND d.item_name <> 'Stock'
              AND COALESCE(d.sub_total, 0) > 0
        ),
        typed_sales AS (
            SELECT
                *,
                CASE
                    WHEN final_order_type IN ('TAKEAWAY', 'DELIVERY')
                        THEN COALESCE(gross_amount, 0) / 1.09
                    ELSE COALESCE(gross_amount, 0)
                END AS net_amount
            FROM sales_data
        )
        SELECT
            COUNT(DISTINCT receipt_no) AS total_transactions,
            COUNT(DISTINCT CASE WHEN final_order_type = 'DINE-IN' THEN receipt_no END)
                AS dine_in_transactions,
            ROUND(SUM(CASE WHEN final_order_type = 'DINE-IN' THEN gross_amount ELSE 0 END), 2)
                AS dine_in_sales,
            ROUND(SUM(CASE WHEN final_order_type = 'DINE-IN' THEN service_charge ELSE 0 END), 2)
                AS dine_in_service_charge,
            COUNT(DISTINCT CASE WHEN final_order_type = 'TAKEAWAY' THEN receipt_no END)
                AS takeaway_transactions,
            ROUND(SUM(CASE WHEN final_order_type = 'TAKEAWAY' THEN gross_amount ELSE 0 END), 2)
                AS takeaway_sales,
            ROUND(SUM(CASE WHEN final_order_type = 'TAKEAWAY' THEN net_amount ELSE 0 END), 2)
                AS takeaway_net_sales,
            COUNT(DISTINCT CASE WHEN final_order_type = 'DELIVERY' THEN receipt_no END)
                AS delivery_transactions,
            ROUND(SUM(CASE WHEN final_order_type = 'DELIVERY' THEN gross_amount ELSE 0 END), 2)
                AS delivery_sales,
            ROUND(SUM(CASE WHEN final_order_type = 'DELIVERY' THEN net_amount ELSE 0 END), 2)
                AS delivery_net_sales,
            COUNT(DISTINCT CASE WHEN final_order_type = 'UNKNOWN' THEN receipt_no END)
                AS unknown_transactions,
            ROUND(SUM(CASE WHEN final_order_type = 'UNKNOWN' THEN gross_amount ELSE 0 END), 2)
                AS unknown_sales,
            ROUND(SUM(gross_amount), 2) AS total_sales,
            ROUND(SUM(net_amount), 2) AS total_net_sales,
            ROUND(SUM(item_tax), 2) AS total_tax,
            MIN(receipt_date) AS first_date,
            MAX(receipt_date) AS last_date,
            COUNT(DISTINCT receipt_no) AS unique_receipts,
            STRING_AGG(DISTINCT final_order_type, ', ') AS order_types_found
        FROM typed_sales
        """
        
        # Execute the query
        sales_df = conn.execute(query).fetchdf()
        
        # Add store code to the results
        sales_df['store_code'] = store_code
        
        return sales_df
        
    except Exception as e:
        print(f"Error calculating sales: {str(e)}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        if 'conn' in locals():
            conn.close()

def main():
    print("\n=== Calculating Sales for Store 204 ===\n")
    
    # Calculate sales for store 204
    store_code = '204'
    sales_df = calculate_store_sales(store_code)
    
    if sales_df is not None and not sales_df.empty:
        print("\n=== Sales Summary ===")
        print(sales_df.to_string(index=False))
        
        # Print a more readable summary
        print("\n=== Detailed Summary ===")
        print(f"Dine-in Transactions: {sales_df['dine_in_transactions'].iloc[0]}")
        print(f"Dine-in Sales: ${sales_df['dine_in_sales'].iloc[0]:.2f}")
        print(f"\nTakeaway Transactions: {sales_df['takeaway_transactions'].iloc[0]}")
        print(f"Takeaway Sales: ${sales_df['takeaway_sales'].iloc[0]:.2f}")
        print(f"\nTotal Sales: ${sales_df['total_sales'].iloc[0]:.2f}")
        print(f"Total GST: ${sales_df['total_tax'].iloc[0]:.2f}")
        print(f"\nDate Range: {sales_df['first_date'].iloc[0]} to {sales_df['last_date'].iloc[0]}")
    else:
        print("\nNo sales data found or error occurred.")
        
        # Try to get some basic info about the data
        try:
            db_path = ".pos_cache/sushi_epoint_pos_live/pos_20240102.duckdb"
            conn = duckdb.connect(database=db_path, read_only=True)
            
            # Check store codes
            print("\nChecking available store codes in raw_receipts:")
            stores = conn.execute("""
                SELECT store_code, COUNT(*) as receipt_count 
                FROM raw_receipts 
                GROUP BY store_code 
                ORDER BY receipt_count DESC
            """).fetchdf()
            print(stores.head(10))
            
            # Check data for store 204
            print("\nSample data for store 204:")
            sample = conn.execute("""
                SELECT d.sales_no, d.item_name, d.sub_total, d.tax_amt, d.take_away_item, r.terminal, r.date, r.time
                FROM raw_details d
                JOIN raw_receipts r ON d.sales_no = r.receipt_no
                WHERE d.store_code = '204'
                AND d.item_name != 'Stock'
                AND d.sub_total > 0
                LIMIT 5
            """).fetchdf()
            print(sample)
            
            conn.close()
        except Exception as e:
            print(f"Error checking database: {str(e)}")
            print(f"Error calculating sales: {str(e)}")

if __name__ == "__main__":
    main()
