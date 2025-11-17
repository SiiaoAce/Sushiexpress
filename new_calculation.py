#!/usr/bin/env python3
"""Final outlet summary with proper store name mapping"""

import pandas as pd
from pathlib import Path

# Store name mapping
STORE_NAMES = {
    "204": "204- WEST GATE",
    "206": "206- SELETAR MALL",
    "207": "207- SERANGOON NEX",
    "208": "208- SUN PLAZA",
    "209": "209- IMM MALL",
    "211": "211- WHITE SANDS",
    "212": "212- JURONG POINT",
    "214": "214- HILLION MALL",
    "215": "215- HEARTLAND MALL",
    "216": "216- WATERWAY POINT",
    "217": "217- HEARTBEAT BEDOK",
    "218": "218- NORTHPOINT CITY",
    "221": "221- FUNAN",
    "222": "222- PAYA LEBAR QUARTER",
    "223": "223- HOUGANG MALL",
    "224": "224- PARKWAY PARADE",
    "225": "225- CLEMENTI MALL",
    "226": "226- CENTURY SQUARE",
    "227": "227- SENGKANG GRAND",
    "301": "301 - WOODLAND",
    "302": "302 - GRANTRAL MALL",
    "303": "303 - TAMPINES",
    "304": "304 - TOA PAYOH",
    "306": "306 - J8",
    "307": "307 - HGTO",
    "308": "308 - OASIS",
    "309": "309 - YEW TEE SQUARE",
    "310": "310 - SENGKANG MRT",
    "311": "311 - THE POIZ CENTRE",
    "312": "312 - ANG MO KIO",
    "313": "313 - CANBERRA PLAZA",
    "314": "314 - BUKIT GOMBAK",
    "401": "401 - BUGIS JUNCTION",
    "402": "402 - 313 SOMERSET",
}

def main():
    # Load the specific 2024-01-01 parquet file
    parquet_path = Path(".pos_cache/sushi_epoint_pos_live/2024/01/01/pos_20240101_details.parquet")
    
    if not parquet_path.exists():
        print(f"Parquet file not found: {parquet_path}")
        return
    
    df = pd.read_parquet(parquet_path)
    print(f"Loaded {parquet_path}: {len(df)} records")
    
    # Filter out void receipts, wastage items, and empty store codes
    df = df[(df['is_void'] == 0) & (df['is_wastage'] == 0) & (df['store_code3'].notna()) & (df['store_code3'] != '')]
    print(f"After filtering voids and wastage: {len(df)} records")
    
    # Filter out receipts with staff products and zero total amount
    # First identify receipts with staff products
    staff_receipts = df[df['item_name'].str.contains('STAFF', case=False, na=False)]['sales_no'].unique()
    print(f"Found {len(staff_receipts)} receipts with staff products")
    
    # Calculate total amount per receipt
    receipt_totals = df.groupby('sales_no')['amount'].sum().reset_index()
    receipt_totals.columns = ['sales_no', 'receipt_total']
    
    # Identify receipts with staff products AND zero total amount
    zero_staff_receipts = receipt_totals[
        (receipt_totals['sales_no'].isin(staff_receipts)) & 
        (receipt_totals['receipt_total'] == 0)
    ]['sales_no'].tolist()
    
    print(f"Found {len(zero_staff_receipts)} receipts with staff products and zero total amount")
    
    # Filter out these receipts
    df = df[~df['sales_no'].isin(zero_staff_receipts)]
    print(f"After filtering staff zero receipts: {len(df)} records")
    
    # Map store names
    df['store_name'] = df['store_code3'].map(STORE_NAMES).fillna(df['store_name'])
    
    # Calculate outlet totals
    outlet_summary = df.groupby(['store_code3', 'store_name']).agg({
        'amount': 'sum',
        'net_amount': 'sum',
        'sales_no': 'nunique'
    }).reset_index()
    
    # Calculate service charge
    service_charge = df[df['item_name'] == 'SERVICE CHARGE 10%'].groupby(['store_code3', 'store_name'])['amount'].sum().reset_index()
    service_charge.columns = ['store_code3', 'store_name', 'service_charge']
    
    # Merge service charge
    outlet_summary = outlet_summary.merge(service_charge, on=['store_code3', 'store_name'], how='left')
    outlet_summary['service_charge'] = outlet_summary['service_charge'].fillna(0)
    
    # Sort by store code
    outlet_summary = outlet_summary.sort_values('store_code3')
    
    # Display results
    print("="*110)
    print(f"OUTLET SALES SUMMARY - 1 January 2024")
    print("="*110)
    print(f"{'Store Code':<12} {'Store Name':<30} {'Gross Sales':>12} {'Nett Sales':>12} {'Service Charge':>15} {'Transactions':>12}")
    print("-"*110)
    
    total_gross = 0
    total_nett = 0
    total_svc = 0
    total_tx = 0
    
    for _, row in outlet_summary.iterrows():
        total_gross += row['amount']
        total_nett += row['net_amount']
        total_svc += row['service_charge']
        total_tx += row['sales_no']
        
        print(f"{row['store_code3']:<12} {row['store_name'][:30]:<30} ${row['amount']:>11.2f} ${row['net_amount']:>11.2f} ${row['service_charge']:>14.2f} {row['sales_no']:>12}")
    
    print("-"*110)
    print(f"{'TOTAL':<12} {'':30} ${total_gross:>11.2f} ${total_nett:>11.2f} ${total_svc:>14.2f} {total_tx:>12}")
    print("="*110)
    
    # Store 204 details
    print("\nSTORE 204 - DETAILED BREAKDOWN:")
    print("="*50)
    store_204 = df[(df['store_code3'] == '204') & (df['is_void'] == 0)]
    
    # Sales category breakdown
    category_breakdown = store_204.groupby('sales_category').agg({
        'amount': 'sum',
        'net_amount': 'sum',
        'sales_no': 'nunique'
    }).reset_index()
    
    print("\nSales Category Breakdown:")
    for _, row in category_breakdown.iterrows():
        print(f"  {row['sales_category']:<12} Gross: ${row['amount']:>9.2f} -> Nett: ${row['net_amount']:>9.2f} ({row['sales_no']} receipts)")
    
    # Service charge verification
    svc_204 = store_204[store_204['item_name'] == 'SERVICE CHARGE 10%']
    print(f"\nService Charge Summary:")
    print(f"  Total service charge transactions: {len(svc_204)}")
    print(f"  Total service charge amount: ${svc_204['amount'].sum():.2f}")
    
    # GST check
    gst_204 = store_204[store_204['item_name'] == 'GST 9%']
    print(f"  GST items found: {len(gst_204)} (should be 0)")
    
    # Void check
    void_204 = df[(df['store_code3'] == '204') & (df['is_void'] == 1)]
    print(f"  Void receipts: {len(void_204)} (should be 0)")

if __name__ == "__main__":
    main()
