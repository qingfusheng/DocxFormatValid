# -*- coding: gbk -*-
import requests


def convert_docx_to_pdf(docx_file_path):
    url = "http://127.0.0.1:5000/upload/"  # �޸�Ϊ FastAPI �������ĵ�ַ�Ͷ˿ں�
    files = {"file": open(docx_file_path, "rb")}
    response = requests.post(url, files=files)

    if response.status_code == 200:
        pdf_data = response.content
        with open("output.pdf", "wb") as f:
            f.write(pdf_data)
        print("Conversion successful. The PDF has been saved as 'output.pdf'.")
    else:
        print(f"Conversion failed. Status code: {response.status_code}")
        print(response.text)


if __name__ == "__main__":
    docx_file_path = r"C:\code\runtime\Debug\Papers\Ф¶¶��ҵ����.docx"  # �޸�Ϊ�㱾�ص� .docx �ļ�·��
    convert_docx_to_pdf(docx_file_path)
