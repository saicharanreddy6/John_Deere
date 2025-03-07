import os
import sys
import logging
import argparse
from bs4 import BeautifulSoup
from datetime import datetime

# Set up script paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(SCRIPT_DIR)  # Add the script directory to the Python path
DATA_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "data"))
LOG_DIR = os.path.join(SCRIPT_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

# Import utility functions from utils.py
from utils import (
    process_section, process_block, process_block_head, process_block_notices,
    process_block_figures, process_block_paragraphs, process_block_lists,
    process_block_tables, process_block_subblocks, extract_text,
    markdown_to_csv, markdown_to_excel
)

# Set up logging
log_filename = os.path.join(LOG_DIR, f"xml_processor_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("xml_processor")

def xml_to_markdown_conversion(xml_file_path):
    """
    Converting an XML document to Markdown format using BeautifulSoup.
    """
    try:
        with open(xml_file_path, 'r', encoding='utf-8') as file:
            xml_content = file.read()

        soup = BeautifulSoup(xml_content, 'xml')
        md_text = ""
        pubinfo = soup.find('pubinfo')
        if pubinfo:
            pubtitle = pubinfo.find('pubtitle')
            if pubtitle and pubtitle.text:
                md_text += f"# {pubtitle.text.strip()}\n\n"
            edition = pubinfo.find('edition')
            if edition and edition.text:
                md_text += f"**{edition.text.strip()}**\n\n"
            litho = pubinfo.find('litho')
            if litho and litho.text:
                md_text += f"**{litho.text.strip()}**\n\n"

        intro = soup.find('intro')
        if intro:
            for block in intro.find_all('block', recursive=False):
                md_text += process_block(block)

        for section in soup.find_all('omsection'):
            md_text += process_section(section)

        return md_text

    except Exception as e:
        logger.error(f"Error during Markdown conversion: {str(e)}")
        raise

def main():
    parser = argparse.ArgumentParser(
        description="Convert XML to Markdown, then parse Markdown tables to CSV and Excel."
    )
    parser.add_argument(
        "xml_file",
        nargs="?",
        default="data/test.xml",
        help="Path to the XML file to convert (default: data/test.xml)"
    )
    parser.add_argument(
        "output_file",
        nargs="?",
        default="data/output.md",
        help="Path to save the Markdown output (default: data/output.md)"
    )
    args = parser.parse_args()

    try:
        # 1) Convert XML -> Markdown
        markdown_output = xml_to_markdown_conversion(args.xml_file)
        with open(args.output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_output)
        logger.info(f"Markdown conversion successful. Output saved to {args.output_file}")
        print(f"Markdown conversion successful. Output saved to {args.output_file}")

        # 2) Convert Markdown -> CSV
        csv_file_path = args.output_file.replace('.md', '.csv')
        markdown_to_csv(args.output_file, csv_file_path)
        logger.info(f"CSV conversion successful. Output saved to {csv_file_path}")
        print(f"CSV conversion successful. Output saved to {csv_file_path}")

        # 3) Convert Markdown -> Excel with formatting.
        excel_file_path = args.output_file.replace('.md', '.xlsx')
        markdown_to_excel(args.output_file, excel_file_path)
        logger.info(f"Excel conversion successful. Output saved to {excel_file_path}")
        print(f"Excel conversion successful. Output saved to {excel_file_path}")

    except Exception as e:
        logger.error(f"Error: {e}")
        print(f"Error: {e}")

if __name__ == "__main__":
    main()