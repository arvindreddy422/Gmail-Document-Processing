from agent import run_agent
from extract_pdf_md import process_input_folder
from json_from_md import process_md_folder


def main():
    run_agent()
    process_input_folder()
    process_md_folder()

main()