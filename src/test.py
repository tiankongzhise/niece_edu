import pandas
import os


def get_pdf_file_name(file_path):
    all_file_name = os.listdir(file_path)
    pdf_file_name = [file for file in all_file_name if file.endswith('.pdf')]
    return pdf_file_name


if __name__ == '__main__':
    data = get_pdf_file_name('./data/doc/')
    print(data)