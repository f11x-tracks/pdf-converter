#!/usr/bin/env python3
"""
PDF to Excel Converter (Single Sheet Mode)
Converts PDF files containing tables to Excel format using multiple extraction methods.
All extracted tables are combined into a single Excel sheet.
"""

import pandas as pd
import pdfplumber
import tabula
from pathlib import Path
import sys
import logging
from typing import List, Optional
import argparse

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PDFToExcelConverter:
    def __init__(self, pdf_path: str, output_path: Optional[str] = None):
        """
        Initialize the PDF to Excel converter.
        
        Args:
            pdf_path (str): Path to the input PDF file
            output_path (str, optional): Path for the output Excel file
        """
        self.pdf_path = Path(pdf_path)
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        if output_path:
            self.output_path = Path(output_path)
        else:
            self.output_path = self.pdf_path.with_suffix('.xlsx')
    
    def extract_tables_with_pdfplumber(self) -> List[pd.DataFrame]:
        """
        Extract tables using pdfplumber library.
        
        Returns:
            List[pd.DataFrame]: List of extracted tables as DataFrames
        """
        tables = []
        logger.info("Extracting tables using pdfplumber...")
        
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    logger.info(f"Processing page {page_num}/{len(pdf.pages)}")
                    
                    # Extract tables from the page
                    page_tables = page.extract_tables()
                    
                    for table_num, table in enumerate(page_tables, 1):
                        if table and len(table) > 1:  # Ensure table has data
                            try:
                                # Convert to DataFrame
                                df = pd.DataFrame(table[1:], columns=table[0])
                                df = self._clean_dataframe(df)
                                
                                if not df.empty:
                                    df.name = f"Page_{page_num}_Table_{table_num}"
                                    tables.append(df)
                                    logger.info(f"Extracted table from page {page_num}, table {table_num}: {df.shape}")
                            except Exception as e:
                                logger.warning(f"Error processing table on page {page_num}: {e}")
                    
                    # Also try to extract text and look for structured data
                    if not page_tables:
                        text = page.extract_text()
                        if text and self._looks_like_tabular_data(text):
                            try:
                                df = self._extract_table_from_text(text, page_num)
                                if df is not None and not df.empty:
                                    df.name = f"Page_{page_num}_Text_Table"
                                    tables.append(df)
                                    logger.info(f"Extracted text table from page {page_num}: {df.shape}")
                            except Exception as e:
                                logger.warning(f"Error extracting text table from page {page_num}: {e}")
        
        except Exception as e:
            logger.error(f"Error extracting tables with pdfplumber: {e}")
        
        return tables
    
    def extract_tables_with_tabula(self) -> List[pd.DataFrame]:
        """
        Extract tables using tabula-py library.
        
        Returns:
            List[pd.DataFrame]: List of extracted tables as DataFrames
        """
        tables = []
        logger.info("Extracting tables using tabula-py...")
        
        try:
            # Extract all tables from all pages
            dfs = tabula.read_pdf(str(self.pdf_path), pages='all', multiple_tables=True)
            
            for i, df in enumerate(dfs, 1):
                if not df.empty:
                    df = self._clean_dataframe(df)
                    if not df.empty:
                        df.name = f"Tabula_Table_{i}"
                        tables.append(df)
                        logger.info(f"Extracted table {i} with tabula: {df.shape}")
        
        except Exception as e:
            logger.error(f"Error extracting tables with tabula: {e}")
        
        return tables
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Clean and preprocess the extracted DataFrame.
        
        Args:
            df (pd.DataFrame): Raw extracted DataFrame
            
        Returns:
            pd.DataFrame: Cleaned DataFrame
        """
        # Remove completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Remove rows where all values are None or empty strings
        df = df[~df.apply(lambda row: all(pd.isna(val) or val == '' for val in row), axis=1)]
        
        # Strip whitespace from string columns
        for col in df.columns:
            try:
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).str.strip()
                    df[col] = df[col].replace('nan', '')
            except Exception:
                # Handle cases where dtype check might fail
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace('nan', '')
        
        return df
    
    def _looks_like_tabular_data(self, text: str) -> bool:
        """
        Check if text contains patterns that suggest tabular data.
        Enhanced to detect various table formats and patterns.
        
        Args:
            text (str): Text to analyze
            
        Returns:
            bool: True if text appears to contain tabular data
        """
        lines = text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        if len(lines) < 3:
            return False
        
        # Enhanced detection patterns
        table_indicators = 0
        
        # Check for consistent column patterns
        separator_count = 0
        for line in lines[:10]:  # Check first 10 lines
            if len(line.split()) > 2:  # Multiple columns
                separator_count += 1
        
        if separator_count > 2:
            table_indicators += 1
        
        # Look for common table separators
        separator_chars = ['|', '\t', '  ', ',']
        for char in separator_chars:
            if sum(1 for line in lines[:5] if char in line) >= 3:
                table_indicators += 1
                break
        
        # Check for numeric patterns (common in tables)
        numeric_lines = 0
        for line in lines[:10]:
            if any(char.isdigit() for char in line):
                numeric_lines += 1
        
        if numeric_lines >= 3:
            table_indicators += 1
        
        # Look for header-like patterns (all caps, underscores, etc.)
        if lines:
            first_line = lines[0]
            if (first_line.isupper() or '_' in first_line or 
                any(word in first_line.lower() for word in ['name', 'date', 'id', 'code', 'amount', 'total'])):
                table_indicators += 1
        
        return table_indicators >= 2
    
    def extract_mixed_content(self) -> tuple[List[pd.DataFrame], List[dict]]:
        """
        Extract tables from PDF (for compatibility, but only returns tables).
        
        Returns:
            tuple: (tables_list, empty_list) - Only tables are extracted for single sheet output
        """
        tables = self.extract_tables_with_pdfplumber()
        return tables, []
    
    def _extract_text_sections(self, text: str, page_num: int) -> List[dict]:
        """
        Simplified text extraction - not used in single sheet mode.
        
        Args:
            text (str): Full page text
            page_num (int): Page number
            
        Returns:
            List[dict]: Empty list (not used in single sheet mode)
        """
        return []
    
    def _classify_text_section(self, text: str) -> str:
        """
        Simplified text classification - not used in single sheet mode.
        
        Args:
            text (str): Text to classify
            
        Returns:
            str: Empty string (not used in single sheet mode)
        """
        return ""
    
    def _extract_table_from_text(self, text: str, page_num: int) -> Optional[pd.DataFrame]:
        """
        Try to extract tabular data from plain text.
        
        Args:
            text (str): Text to parse
            page_num (int): Page number for reference
            
        Returns:
            Optional[pd.DataFrame]: Extracted DataFrame or None
        """
        lines = text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        if len(lines) < 2:
            return None
        
        # Try to identify columns by looking for consistent spacing
        table_data = []
        for line in lines:
            # Split by multiple spaces or tabs
            parts = [part.strip() for part in line.split() if part.strip()]
            if len(parts) > 1:
                table_data.append(parts)
        
        if len(table_data) < 2:
            return None
        
        # Use first row as headers, rest as data
        try:
            max_cols = max(len(row) for row in table_data)
            
            # Pad rows to have the same number of columns
            for row in table_data:
                while len(row) < max_cols:
                    row.append('')
            
            df = pd.DataFrame(table_data[1:], columns=table_data[0])
            return self._clean_dataframe(df)
        except Exception:
            return None
    
    def save_to_excel(self, tables: List[pd.DataFrame]) -> None:
        """
        Save extracted tables to Excel file in a single sheet.
        
        Args:
            tables (List[pd.DataFrame]): List of tables to save
        """
        if not tables:
            logger.warning("No tables found to save")
            return
        
        logger.info(f"Saving {len(tables)} tables to single sheet in {self.output_path}")
        
        self._save_to_single_sheet(tables)
        
        logger.info(f"Excel file saved successfully: {self.output_path}")
    
    def save_mixed_content_to_excel(self, tables: List[pd.DataFrame], text_content: List[dict], 
                                   mixed_format: str = "single_sheet") -> None:
        """
        Save tables to Excel in single sheet format (ignores text content).
        
        Args:
            tables (List[pd.DataFrame]): Extracted tables
            text_content (List[dict]): Ignored in single sheet mode
            mixed_format (str): Ignored - always uses single sheet
        """
        self.save_to_excel(tables)
    
    def _save_to_single_sheet(self, tables: List[pd.DataFrame]) -> None:
        """
        Save all tables to a single Excel sheet with separators.
        
        Args:
            tables (List[pd.DataFrame]): List of tables to save
        """
        combined_rows = []
        
        for i, df in enumerate(tables):
            table_name = getattr(df, 'name', f'Table_{i+1}')
            
            # Add table header
            combined_rows.append([f"=== {table_name} ==="])
            
            # Add the table data
            if not df.empty:
                # Add column headers
                combined_rows.append(list(df.columns))
                
                # Add data rows
                for _, row in df.iterrows():
                    combined_rows.append(list(row))
            
            # Add empty separator row (except after last table)
            if i < len(tables) - 1:
                combined_rows.append([""])
        
        # Create DataFrame from all rows
        if combined_rows:
            # Find the maximum number of columns needed
            max_cols = max(len(row) for row in combined_rows)
            
            # Pad all rows to have the same number of columns
            for row in combined_rows:
                while len(row) < max_cols:
                    row.append('')
            
            # Create column names
            columns = [f'Column_{i+1}' for i in range(max_cols)]
            
            final_df = pd.DataFrame(combined_rows, columns=columns)
            
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name='All_Tables', index=False)
                logger.info(f"Saved all {len(tables)} tables to single sheet: All_Tables")
    
    def _save_to_multiple_sheets(self, tables: List[pd.DataFrame]) -> None:
        """
        Not used in single sheet mode - redirects to single sheet.
        
        Args:
            tables (List[pd.DataFrame]): List of tables to save
        """
        self._save_to_single_sheet(tables)
    
    def convert(self, method: str = 'pdfplumber') -> None:
        """
        Convert PDF to Excel using specified method (always single sheet).
        
        Args:
            method (str): Extraction method ('pdfplumber', 'tabula', or 'both')
        """
        all_tables = []
        
        if method in ['pdfplumber', 'both']:
            pdfplumber_tables = self.extract_tables_with_pdfplumber()
            all_tables.extend(pdfplumber_tables)
        
        # Skip tabula if Java is not available
        if method in ['tabula', 'both']:
            try:
                tabula_tables = self.extract_tables_with_tabula()
                all_tables.extend(tabula_tables)
            except Exception as e:
                logger.warning(f"Tabula extraction failed (Java may not be installed): {e}")
        
        if all_tables:
            self.save_to_excel(all_tables)
            print(f"✅ Successfully converted '{self.pdf_path}' to '{self.output_path}'")
            print(f"📊 Extracted {len(all_tables)} tables")
            print("📋 All tables saved to a single sheet")
        else:
            print(f"❌ No tables found in '{self.pdf_path}'")
            print("💡 This PDF might not contain structured tables, or the content might be in image format.")
            logger.warning("No tables were extracted from the PDF")

def main():
    """Main function to run the PDF to Excel converter in single sheet mode."""
    parser = argparse.ArgumentParser(description='Convert PDF tables to Excel format (single sheet only)')
    parser.add_argument('pdf_file', help='Path to the input PDF file')
    parser.add_argument('-o', '--output', help='Output Excel file path (optional)')
    parser.add_argument('-m', '--method', choices=['pdfplumber', 'tabula', 'both'], 
                        default='both', help='Extraction method to use')
    
    args = parser.parse_args()
    
    try:
        converter = PDFToExcelConverter(args.pdf_file, args.output)
        converter.convert(args.method)
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # If no command line arguments, convert the ProZ-FP.pdf file in single sheet mode
    if len(sys.argv) == 1:
        print("Converting ProZ-FP.pdf to Excel (single sheet mode)...")
        print("This will extract all tables and combine them into one Excel sheet.")
        
        try:
            converter = PDFToExcelConverter("ProZ-FP.pdf")
            converter.convert()
                
        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)
    else:
        main()