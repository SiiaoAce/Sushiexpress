import pandas as pd
import re
from datetime import datetime, date
from pathlib import Path
import duckdb
from typing import Optional, Dict, List, Tuple

class EpointParser:
    def __init__(self):
        self.current_receipt = None
        self.current_details = []
        self.receipts = []
        self.details = []
        self.wastage_records = []
        self.stock_master = None
        self.stock_category_map = {}
        self.payments = []
        self.store_map = {
            '204': 'WEST GATE',
            '206': 'SELETAR MALL',
            '207': 'SERANGOON NEX',
            '208': 'SUN PLAZA',
            '209': 'IMM MALL',
            '211': 'WHITE SANDS',
            '212': 'JURONG POINT',
            '214': 'HILLION MALL',
            '215': 'HEARTLAND MALL',
            '216': 'WATERWAY POINT',
            '217': 'Heartbeat Bedok',
            '218': 'NORTHPOINT CITY',
            '221': 'FUNAN',
            '222': 'PAYA LEBAR QUARTER',
            '225': 'CLEMENTI MALL',
            '226': 'CENTURY SQUARE',
            '227': 'SENGKANG GRAND',
            '228': 'WEST MALL',
            '229': 'PARKWAY PARADE 2',
            '301': 'WOODLAND',
            '302': 'GRANTRAL MALL',
            '303': 'TAMPINES',
            '304': 'TOA PAYOH',
            '306': 'J8',
            '307': 'HGTO',
            '308': 'OASIS',
            '309': 'YEW TEE SQUARE',
            '310': 'SENGKANG MRT',
            '311': 'THE POIZ CENTRE',
            '312': 'ANG MO KIO',
            '313': 'CANBERRA PLAZA',
            '314': 'BT GOMBAK',
            '316': 'PASIR RIS MRT',
            '317': 'NORTHPOINT TO',
            '401': 'Bugis Junction',
            '402': '313 Somerset',
            '403': 'Tampines One',
        }
        
    def load_stock_master(self, stock_master_path: Path):
        """Load stock master data from Excel file"""
        try:
            sample = pd.read_excel(stock_master_path, nrows=1)
            available_cols = [c.strip() if isinstance(c, str) else c for c in sample.columns]
            print(f"Available columns in stock master: {available_cols}")

            def find_column(candidates: List[str]) -> Optional[str]:
                for cand in candidates:
                    for col in available_cols:
                        if isinstance(col, str) and cand.lower() == col.lower():
                            return col
                        if isinstance(col, str) and cand.lower() in col.lower():
                            return col
                return None

            stock_col = find_column(['STOCK', 'Stock code'])
            category_col = find_column(['DEPT', 'Category'])
            sales_cat_col = find_column(['AC_GROUP', 'Sales Category', '業績計算分類'])
            desc_col = find_column(['DESCRIP1', 'Description'])

            required = {
                'Stock code': stock_col,
                'Category': category_col,
                'Sales Category': sales_cat_col
            }

            if any(val is None for val in required.values()):
                print("Warning: Could not identify required columns in stock master. Stock data will not be applied.")
                self.stock_master = None
                return

            use_columns = [val for val in required.values() if val is not None]
            if desc_col:
                use_columns.append(desc_col)

            df = pd.read_excel(stock_master_path, usecols=use_columns, dtype={required['Stock code']: str})

            rename_map = {
                required['Stock code']: 'Stock code',
                required['Category']: 'Category',
                required['Sales Category']: 'Sales Category'
            }
            if desc_col:
                rename_map[desc_col] = 'Description'

            df = df.rename(columns=rename_map)

            for col in ['Stock code', 'Category', 'Sales Category', 'Description']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()

            df = df.drop_duplicates('Stock code')
            self.stock_master = df
            if 'Sales Category' in df.columns:
                self.stock_category_map = dict(zip(df['Stock code'], df['Sales Category']))
            else:
                self.stock_category_map = {}
            print(f"Loaded {len(self.stock_master)} items from stock master")

        except Exception as e:
            print(f"Error loading stock master: {e}")
            import traceback
            traceback.print_exc()
            self.stock_master = None
    
    def parse_excel(self, input_file: Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Parse the Epoint Excel file and return receipts and details DataFrames"""
        print(f"Reading Excel file: {input_file}")
        
        # Read the Excel file without headers
        df = pd.read_excel(input_file, header=None)
        
        # Reset state
        self.current_receipt = None
        self.current_details = []
        self.receipts = []
        self.details = []
        self.wastage_records = []
        self.payments = []
        
        # Process each row
        for idx, row in df.iterrows():
            row_data = row.dropna()
            if len(row_data) == 0:
                continue
                
            # Convert to list and handle NaN values
            row_values = [val if pd.notna(val) else '' for val in row]
            
            # Check for new receipt
            receipt_match = re.match(r'(\d+)-([A-Z]\d+)', str(row_values[0]))
            if receipt_match:
                self._save_current_receipt()
                self._start_new_receipt(receipt_match.group(0), row_values, df, idx)
            
            # Process the current receipt
            if self.current_receipt:
                self._process_receipt_row(row_values)
        
        # Save the last receipt
        self._save_current_receipt()
        
        receipts_df = pd.DataFrame(self.receipts) if self.receipts else self._create_empty_receipts_df()
        details_df = pd.DataFrame(self.details) if self.details else self._create_empty_details_df()
        payments_df = pd.DataFrame(self.payments, columns=['c_date','store_name','sales_no','payment_name','tender_amt']) if self.payments else pd.DataFrame(columns=['c_date','store_name','sales_no','payment_name','tender_amt'])
        
        # Add stock master data to details
        if self.stock_master is not None and not details_df.empty:
            details_df = details_df.merge(
                self.stock_master,
                left_on='item_code',
                right_on='Stock code',
                how='left'
            )
            
            details_df['category_code'] = details_df.get('Category')
            details_df['sales_category'] = details_df.get('Sales Category')
            details_df['take_away_item'] = details_df['category_code'].apply(lambda s: 'Y' if isinstance(s, str) and 'TAKEAWAY' in s.upper() else 'N')
        
        return receipts_df, details_df, payments_df
    
    def _start_new_receipt(self, receipt_no: str, row_values: list, df: pd.DataFrame, idx: int):
        """Initialize a new receipt"""
        # Get the next row which contains date, time, etc.
        next_row = []
        if idx + 1 < len(df):
            next_row = df.iloc[idx + 1].fillna('').tolist()
        
        # Extract receipt info from surrounding rows
        receipt_date = None
        receipt_time = ''
        terminal = ''
        cashier = ''

        context_rows = [row_values]
        if next_row:
            context_rows.append(next_row)

        def try_parse_datetime(value):
            if isinstance(value, (datetime, pd.Timestamp)):
                return pd.to_datetime(value)
            if isinstance(value, str):
                val = value.strip()
                if not val:
                    return None
                try:
                    return pd.to_datetime(val)
                except Exception:
                    return None
            return None

        for row in context_rows:
            for val in row:
                if receipt_date is None:
                    dt_val = try_parse_datetime(val)
                    if dt_val is not None:
                        receipt_date = dt_val.date()
                        if dt_val.time() != datetime.min.time():
                            receipt_time = dt_val.strftime('%H:%M:%S')
                if not receipt_time and isinstance(val, str):
                    time_match = re.search(r'(\d{1,2}:\d{2}:\d{2})', val)
                    if time_match:
                        receipt_time = time_match.group(1)
                if not terminal and isinstance(val, str):
                    term_match = re.search(r'(T\d{3})', val.upper())
                    if term_match:
                        terminal = term_match.group(1)
                if not cashier and isinstance(val, str):
                    trimmed = val.strip()
                    if trimmed and trimmed.isdigit() and len(trimmed) >= 3:
                        cashier = trimmed
        
        if not receipt_time and receipt_date is not None:
            receipt_time = '00:00:00'
        
        store_code = receipt_no.split('-')[0]
        store_name = self.store_map.get(store_code, store_code)
        self.current_receipt = {
            'receipt_no': receipt_no,
            'store_code': store_code,
            'store_name': store_name,
            'date': receipt_date,
            'time': receipt_time,
            'terminal': terminal,
            'cashier': cashier,
            'total_amount': 0,
            'payment_type': 'UNKNOWN',
            'payment_amount': 0,
            'is_void': 0,
            'is_wastage': 0,
            'is_term': 0,
            'has_staff': 0,
            'service_charge_total': 0.0,
            'data_source': 'epoint',
            'brand': 'SE'  # Default to Sushi Express
        }
    
    def _process_receipt_row(self, row_values: list):
        """Process a row within a receipt"""
        if not self.current_receipt:
            return
            
        # Check for void indicators anywhere in the row
        row_text = [str(val).upper() for val in row_values if isinstance(val, str)]
        if any('TRANVOID' in text or text.strip() == 'VOID' for text in row_text):
            self.current_receipt['is_void'] = 1
        
        # Check for item line (starts with stock code)
        stock_code_match = re.match(r'^[A-Za-z0-9_]+$', str(row_values[0]))
        if stock_code_match and len(row_values) >= 10:
            self._process_item(row_values)
        
        # Check for service charge
        if len(row_values) >= 3 and 'SERVICE CHARGE' in str(row_values[2]).upper():
            self._accumulate_service_charge(row_values)
        
        # Check for GST
        elif len(row_values) >= 3 and 'GST' in str(row_values[2]).upper():
            self._process_tax(row_values)
        
        # Check for payment methods
        elif len(row_values) >= 3 and any(pmt in str(row_values[2]).upper() 
                                       for pmt in ['VISA', 'MASTER', 'AMEX', 'CASH', 'UNIONPAY', 'NETS', 'DBS', 'MAX', 'GRAB', 'VOUCHER']):
            self._process_payment(row_values)
        
        # Check for wastage
        elif len(row_values) >= 2 and 'WASTAGE' in str(row_values[1]).upper():
            self.current_receipt['is_wastage'] = 1
        
        # Check for receipt total
        elif len(row_values) >= 8 and 'RCP. AMOUNT' in str(row_values[6]).upper():
            self._process_receipt_total(row_values)

        # Check for Term in Column F
        if len(row_values) > 5 and isinstance(row_values[5], str) and row_values[5].strip().upper().startswith('TERM T00'):
            self.current_receipt['is_term'] = 1
    
    def _process_item(self, row_values: list):
        """Process an item line"""
        try:
            item_code = str(row_values[0])
            item_name = str(row_values[2]) if len(row_values) > 2 else ''
            
            # Skip if marked as void in Column R (index 17)
            if len(row_values) > 17 and row_values[17]:
                return
            
            # Get quantity (Column F, index 5)
            qty = self._parse_float(row_values[5]) if len(row_values) > 5 else 0
            
            # Get amount (Column I, index 8) and tax (Column J, index 9)
            amount = self._parse_float(row_values[8]) if len(row_values) > 8 else 0
            tax = self._parse_float(row_values[9]) if len(row_values) > 9 else 0
            
            # Check for discount
            is_discount = '$DISC' in item_name.upper()
            if 'STAFF' in item_name.upper():
                self.current_receipt['has_staff'] = 1
            
            base_datetime = pd.to_datetime(f"{self.current_receipt['date']} {self.current_receipt['time']}", errors='coerce')
            base_detail = {
                'c_date': base_datetime,
                'store_code': self.current_receipt['store_code'],
                'store_name': self.current_receipt['store_name'],
                'receipt_no': self.current_receipt['receipt_no'],
                'sales_no': self.current_receipt['receipt_no'],
                'item_code': item_code,
                'item_name': item_name,
                'category_code': '',
                'qty': qty,
                'sub_total': amount,
                'disc_name': '',
                'disc_amt': 0.0,
                'pro_disc_amt': 0.0,
                'svc_amt': 0.0,
                'tax_amt': tax,
                'take_away_item': 'N',
                'item_sub_total': amount,
                'payment_methods': '',
                'amount': amount,
                'tax': tax,
                'is_void': 1 if self.current_receipt.get('is_void') else 0,
                'is_discount': 0,
                'is_tax': 0,
                'is_service_charge': 0,
                'sales_category': self._lookup_sales_category(item_code)
            }

            sales_cat = base_detail['sales_category']
            if isinstance(sales_cat, str):
                cat_upper = sales_cat.upper()
                if cat_upper == 'TAKEAWAY':
                    base_detail['take_away_item'] = 'Y'
            
            if is_discount:
                discount_detail = base_detail.copy()
                discount_detail.update({
                    'qty': 1,
                    'sub_total': 0.0,
                    'disc_name': item_name,
                    'disc_amt': -abs(amount),
                    'pro_disc_amt': abs(amount),
                    'item_sub_total': 0.0,
                    'amount': 0.0,
                    'tax': 0.0,
                    'is_discount': 1
                })
                self.current_details.append(discount_detail)
                return

            self.current_details.append(base_detail)
            
        except Exception as e:
            print(f"Error processing item: {e}")

    def _lookup_sales_category(self, item_code: str) -> Optional[str]:
        if not item_code:
            return None
        if not self.stock_category_map:
            return None
        code = str(item_code).strip()
        if not code:
            return None
        if code in self.stock_category_map:
            return self.stock_category_map[code]
        trimmed = code.lstrip('0')
        if trimmed and trimmed in self.stock_category_map:
            return self.stock_category_map[trimmed]
        return None
    
    def _accumulate_service_charge(self, row_values: list):
        try:
            amount = 0.0
            if len(row_values) > 8:
                amount = self._parse_float(row_values[8])

            if amount == 0.0:
                for val in row_values:
                    candidate = 0.0
                    if isinstance(val, (int, float)):
                        candidate = float(val)
                    elif isinstance(val, str):
                        if '%' in val:
                            continue
                        cleaned = val.strip()
                        if not cleaned or any(ch.isalpha() for ch in cleaned):
                            continue
                        candidate = self._parse_float(cleaned)

                    if candidate != 0.0:
                        amount = candidate
                        break

            self.current_receipt['service_charge_total'] = self.current_receipt.get('service_charge_total', 0.0) + amount
        except Exception as e:
            print(f"Error processing service charge: {e}")
    
    def _process_tax(self, row_values: list):
        try:
            tax = self._parse_float(row_values[8]) if len(row_values) > 8 else 0
            self.current_details.append({
                'c_date': pd.to_datetime(f"{self.current_receipt['date']} {self.current_receipt['time']}", errors='coerce'),
                'store_name': self.current_receipt['store_name'],
                'sales_no': self.current_receipt['receipt_no'],
                'item_code': 'GST',
                'item_name': 'GST 9%',
                'category_code': '',
                'qty': 1,
                'sub_total': 0.0,
                'disc_name': '',
                'disc_amt': 0.0,
                'pro_disc_amt': 0.0,
                'svc_amt': 0.0,
                'tax_amt': tax,
                'take_away_item': 'N',
                'item_sub_total': 0.0,
                'payment_methods': ''
            })
        except Exception as e:
            print(f"Error processing tax: {e}")
    
    def _process_payment(self, row_values: list):
        try:
            raw = str(row_values[2])
            payment_type = raw.split(':')[0].strip()
            amount = 0.0
            for val in row_values[2:]:
                try:
                    if isinstance(val, (int, float)):
                        amount = float(val)
                        break
                    elif isinstance(val, str):
                        num_match = re.search(r'([\d.,]+)', val.replace(',', ''))
                        if num_match:
                            amount = float(num_match.group(1))
                            break
                except (ValueError, TypeError):
                    continue
            self.current_receipt['payment_type'] = payment_type
            self.current_receipt['payment_amount'] = amount
            self.current_receipt['total_amount'] = amount
            self.payments.append([
                pd.to_datetime(f"{self.current_receipt['date']} {self.current_receipt['time']}", errors='coerce'),
                self.current_receipt['store_name'],
                self.current_receipt['receipt_no'],
                payment_type,
                amount,
            ])
            for d in self.current_details:
                d['payment_methods'] = payment_type
        except Exception as e:
            print(f"Error processing payment: {e}")
    
    def _process_receipt_total(self, row_values: list):
        """Process receipt total line"""
        try:
            if len(row_values) > 8:
                total = self._parse_float(row_values[7])
                tax = self._parse_float(row_values[8]) if len(row_values) > 8 else 0
                
                self.current_receipt['total_amount'] = total
                self.current_receipt['tax_amount'] = tax
                
        except Exception as e:
            print(f"Error processing receipt total: {e}")
    
    def _save_current_receipt(self):
        """Save the current receipt and its details"""
        if not self.current_receipt:
            return
        if self.current_receipt.get('is_wastage'):
            self.wastage_records.append({
                'receipt_no': self.current_receipt['receipt_no'],
                'date': self.current_receipt['date'],
                'time': self.current_receipt['time'],
                'store_code': self.current_receipt['store_code'],
                'amount': self.current_receipt.get('total_amount', 0),
                'reason': 'WASTAGE'
            })
        if self.current_details:
            if self.current_receipt.get('is_void') == 1:
                self.current_receipt = None
                self.current_details = []
            elif self.current_receipt.get('is_term') == 1:
                self.current_receipt = None
                self.current_details = []
            elif self.current_receipt.get('has_staff') == 1 and float(self.current_receipt.get('total_amount', 0) or 0) == 0.0:
                self.current_receipt = None
                self.current_details = []
            else:
                disc_total = sum(abs(d.get('pro_disc_amt', 0.0)) for d in self.current_details if d.get('disc_name'))
                items = [d for d in self.current_details if not d.get('disc_name') and d.get('item_name') not in ['GST 9%']]
                base_sum = sum(d.get('item_sub_total', 0.0) for d in items)
                for d in items:
                    share = (d.get('item_sub_total', 0.0) / base_sum) if base_sum > 0 else 0.0
                    allocated = disc_total * share
                    d['disc_amt'] = -allocated if allocated > 0 else 0.0
                    d['pro_disc_amt'] = allocated if allocated > 0 else 0.0
                    d['item_sub_total'] = max(d.get('sub_total', 0.0) - allocated, 0.0)
                dine_items = [d for d in items if (d.get('sales_category') or '').upper() == 'DINE-IN']
                dine_sum = sum(d.get('item_sub_total', 0.0) for d in dine_items)
                svc_total = float(self.current_receipt.get('service_charge_total', 0.0) or 0.0)
                for d in dine_items:
                    share = (d.get('item_sub_total', 0.0) / dine_sum) if dine_sum > 0 else 0.0
                    allocated_svc = svc_total * share
                    d['svc_amt'] = allocated_svc
                    # Update per-item totals to include allocated service charge
                    d['item_sub_total'] = d.get('item_sub_total', 0.0) + allocated_svc
                    d['amount'] = d.get('amount', 0.0) + allocated_svc
                    d['sub_total'] = d.get('sub_total', 0.0) + allocated_svc
                self.receipts.append(self.current_receipt)
                self.details.extend(self.current_details)
        self.current_receipt = None
        self.current_details = []
    
    @staticmethod
    def _parse_float(value) -> float:
        """Safely parse a float value"""
        if pd.isna(value) or value == '':
            return 0.0
        try:
            if isinstance(value, str):
                # Remove any non-numeric characters except decimal point and negative sign
                value = re.sub(r'[^\d.-]', '', value)
            return float(value)
        except (ValueError, TypeError):
            return 0.0
    
    @staticmethod
    def _create_empty_receipts_df() -> pd.DataFrame:
        """Create an empty receipts DataFrame with the correct schema"""
        return pd.DataFrame(columns=[
            'receipt_no', 'store_code', 'date', 'time', 'terminal', 'cashier',
            'total_amount', 'payment_type', 'payment_amount', 'is_void', 'is_wastage',
            'data_source', 'brand', 'tax_amount'
        ])
    
    @staticmethod
    def _create_empty_details_df() -> pd.DataFrame:
        """Create an empty details DataFrame with the correct schema"""
        return pd.DataFrame(columns=[
            'receipt_no', 'store_code', 'date', 'time', 'item_code', 'item_name',
            'qty', 'amount', 'tax', 'is_void', 'is_discount', 'is_tax',
            'is_service_charge', 'terminal', 'cashier', 'data_source', 'brand'
        ])


def save_to_duckdb(receipts_df: pd.DataFrame, details_df: pd.DataFrame, payments_df: pd.DataFrame, output_file: Path):
    """Save the parsed data to a DuckDB database"""
    # Create output directory if it doesn't exist
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Connect to DuckDB
    conn = duckdb.connect(str(output_file))
    
    conn.execute("DROP TABLE IF EXISTS raw_receipts")
    conn.execute("DROP TABLE IF EXISTS raw_details")
    conn.execute("DROP TABLE IF EXISTS raw_payments")
    
    # Write data to DuckDB using direct SQL
    if not receipts_df.empty:
        # Register the DataFrame as a temporary table
        conn.register('temp_receipts', receipts_df)
        # Create the table with explicit schema
        conn.execute("""
        CREATE TABLE raw_receipts AS
        SELECT 
            receipt_no,
            store_code,
            date,
            time,
            terminal,
            cashier,
            total_amount,
            payment_type,
            payment_amount,
            is_void,
            is_wastage,
            data_source,
            brand,
            tax_amount
        FROM temp_receipts
        """)
        # Create indexes
        conn.execute("""
        CREATE INDEX idx_receipts_receipt_no ON raw_receipts(receipt_no);
        CREATE INDEX idx_receipts_store_code ON raw_receipts(store_code);
        CREATE INDEX idx_receipts_date ON raw_receipts(date);
        """)
    
    if not details_df.empty:
        conn.register('temp_details', details_df)
        conn.execute("""
        CREATE TABLE raw_details AS
        SELECT
            c_date,
            COALESCE(store_code,'') AS store_code,
            COALESCE(store_name,'') AS store_name,
            COALESCE(sales_no,'') AS sales_no,
            COALESCE(item_code,'') AS item_no,
            COALESCE(item_name,'') AS item_name,
            COALESCE(category_code,'') AS category_code,
            COALESCE(qty, 0.0) AS qty,
            COALESCE(sub_total, 0.0) AS sub_total,
            COALESCE(disc_name,'') AS disc_name,
            COALESCE(disc_amt, 0.0) AS disc_amt,
            COALESCE(pro_disc_amt, 0.0) AS pro_disc_amt,
            COALESCE(svc_amt, 0.0) AS svc_amt,
            COALESCE(tax_amt, 0.0) AS tax_amt,
            COALESCE(take_away_item,'') AS take_away_item,
            COALESCE(item_sub_total, 0.0) AS item_sub_total,
            COALESCE(payment_methods,'') AS payment_methods
        FROM temp_details
        """)
        conn.execute("""
        CREATE INDEX idx_details_sales_no ON raw_details(sales_no);
        CREATE INDEX idx_details_store_name ON raw_details(store_name);
        CREATE INDEX idx_details_c_date ON raw_details(c_date);
        """)

    if not payments_df.empty:
        conn.register('temp_payments', payments_df)
        conn.execute("""
        CREATE TABLE raw_payments AS
        SELECT
            c_date,
            COALESCE(store_name,'') AS store_name,
            COALESCE(sales_no,'') AS sales_no,
            COALESCE(payment_name,'') AS payment_name,
            COALESCE(tender_amt, 0.0) AS tender_amt
        FROM temp_payments
        """)
        conn.execute("""
        CREATE INDEX idx_pay_sales_no ON raw_payments(sales_no);
        CREATE INDEX idx_pay_store_name ON raw_payments(store_name);
        CREATE INDEX idx_pay_c_date ON raw_payments(c_date);
        """)
    
    conn.close()
    print(f"Data saved to {output_file}")


def main():
    # Define paths - use current directory
    script_dir = Path(__file__).parent
    input_file = script_dir / "2 Jan Report.xlsx"
    stock_master_file = script_dir / "Stock master Total.xlsx"
    output_file = script_dir / ".pos_cache" / "sushi_epoint_pos_live" / "pos_20240102.duckdb"
    
    # Create output directory if it doesn't exist
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    print(f"Current directory: {script_dir}")
    print(f"Input file: {input_file}")
    print(f"Stock master: {stock_master_file}")
    print(f"Output file: {output_file}")
    
    # Create output directory if it doesn't exist
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Create parser and load stock master
    parser = EpointParser()
    
    if stock_master_file.exists():
        print(f"Loading stock master from {stock_master_file}")
        parser.load_stock_master(stock_master_file)
    else:
        print(f"Warning: Stock master file not found at {stock_master_file}")
    
    # Continue even if stock master is not available
    if parser.stock_master is None:
        print("Warning: Could not load stock master. Proceeding without it.")
        # Create a dummy stock master with required columns
        parser.stock_master = pd.DataFrame(columns=['Stock code', 'Category Name', 'Sales Category'])
    
    # Parse the Excel file
    print(f"\nParsing {input_file}...")
    receipts_df, details_df, payments_df = parser.parse_excel(input_file)
    
    # Print summary
    print(f"\nParsing complete!")
    print(f"- Receipts: {len(receipts_df)}")
    print(f"- Details: {len(details_df)}")
    
    if not receipts_df.empty:
        print("\nReceipts summary:")
        print(f"Date range: {receipts_df['date'].min()} to {receipts_df['date'].max()}")
        print(f"Stores: {receipts_df['store_code'].nunique()}")
        print(f"Total amount: {receipts_df[receipts_df['is_void'] == 0]['total_amount'].sum():.2f}")
        print("\nPayment methods:")
        print(receipts_df['payment_type'].value_counts())
        
        # Print first few rows for verification
        print("\nSample receipts:")
        print(receipts_df.head().to_string())
    
    if not details_df.empty:
        print("\nSample details:")
        print(details_df.head().to_string())
    
    print("\nSaving to DuckDB...")
    save_to_duckdb(receipts_df, details_df, payments_df, output_file)
    
    # Print wastage records if any
    if parser.wastage_records:
        print("\nWastage records:")
        for record in parser.wastage_records:
            print(f"- {record['receipt_no']} on {record['date']} {record['time']}: {record['amount']:.2f}")


if __name__ == "__main__":
    main()
