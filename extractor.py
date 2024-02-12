import os
import re
import pdfplumber as pb
from win32com.client import Dispatch
import pandas as pd
import sys

class Extractor(object):
    """��ȡ�����ļ�����Ϣ"""

    def __init__(self, file_dir):
        self.file_dir = file_dir
        if os.path.splitext(self.file_dir)[1] in [".doc", ".docx"]:
            try:
                self.__word2pdf()
            except Exception as e:
                print(e)
                return

    def __doc2docx(self):
        """docתΪdocx"""
        w = Dispatch('Word.Application')
        w.Visible = 0
        w.DisplayAlerts = 0
        doc = w.Documents.Open(self.file_dir)
        new_path = os.path.splitext(self.file_dir)[0] + '.docx'
        doc.SaveAs(new_path, 12, False, "", True, "", False, False, False, False)
        doc.Close()
        w.Quit()
        os.remove(self.file_dir)
        self.file_dir = new_path
        return new_path

    def __word2pdf(self):
        """wordתΪpdf"""
        w = Dispatch('Word.Application')
        w.Visible = 0
        w.DisplayAlerts = 0
        doc = w.Documents.Open(self.file_dir)
        new_path = os.path.splitext(self.file_dir)[0] + '.pdf'
        doc.SaveAs(new_path, FileFormat=17)
        doc.Close()
        w.Quit()
        os.remove(self.file_dir)
        self.file_dir = new_path
        return new_path

    def __extract_text(self):
        """��ȡ�ı�����"""
        text = ""
        if os.path.splitext(self.file_dir)[1] == ".pdf":
            pdf = pb.open(self.file_dir)
            for page in pdf.pages:
                text += page.extract_text() if page.extract_text() else ""
        # elif os.path.splitext(self.file_dir)[1] == ".docx":
        #     doc = docx.Document(self.file_dir)
        #     for para in doc.paragraphs:
        #         text += para.text
        return text

    def __extract_words(self):
        """��ȡ����"""
        words = []
        if os.path.splitext(self.file_dir)[1] == ".pdf":
            pdf = pb.open(self.file_dir)
            for page in pdf.pages:
                words += page.extract_words()
        # elif os.path.splitext(self.file_dir)[1] == ".docx":
        #     doc = docx.Document(self.file_dir)
        #     for para in doc.paragraphs:
        #         words.append(para.text)
        return words

    def __search_name(self):
        """��������"""
        names = []
        full_text = self.__extract_text()
        # ��ͨ��"����"�ֶ�ȥ���ҡ�
        for line in full_text.split("\n"):
            if re.search(r"��\s*��", line):
                name = re.findall(r"��\s*��[:��\s]*[\u4e00-\u9fa5]{2,4}", line)[0]
                names.append(re.sub(r"[����:��\s]", "", name))
        # ��"����"�ֶ����Ҳ���������������ֳ���ȥ�²�һ��
        if len(names) < 1:
            for line in re.split(r"\n|\s+", full_text):
                if re.search(r"\d", line):
                    continue
                word = ""
                for w in line:  # ȥ��
                    if w not in word:
                        word += w
                if 2 <= len(word) <= 4:
                    _names = re.findall(r"[\u4e00-\u9fa5]{2,4}", word)
                    names += _names
                    # break
        return names

    def __search_email(self):
        """����Email��ַ"""
        full_words = self.__extract_words()
        email = ""
        for word in full_words:
            if os.path.splitext(self.file_dir)[1] == ".pdf":
                text = word["text"]
            else:
                text = word
            if "@" in text and "." in text:
                for e in re.findall(r"[a-zA-Z0-9_\-.@]+", text):
                    if "@" in e:
                        email = e
                        break
            if email != "":
                break
        return email

    def __search_phone(self):
        """�����绰����"""
        full_text = self.__extract_text()
        phone = ""
        # ֱ��ͨ���ļ�������
        file_name = re.split(r"/+|\\+", self.file_dir)[-1]
        number = re.findall(r"\d{11,13}", file_name)
        if len(number) > 0 and re.search(r"^1", number[0]):
            phone = number[0]
        else:
            # ͨ���ؼ��ʲ���
            for line in re.split(r"[\n\s]+", full_text):
                if "�绰" in line or "�ֻ�" in line:
                    line = re.sub(r"[()������:+\-]", "", line)
                    number = re.findall(r"\d{11,13}", line)[0]
                    phone = re.sub(r"^(86)", "", number)
                    break
            # ֱ��ͨ�����ֳ��Ȳ���
            if phone == "":
                text = re.sub(r"[()����+\-]", "", full_text)
                phones = re.findall(r"\d{11,13}", text)
                phones = [re.sub(r"^(86)", "", p) for p in phones if re.search(r"^1", re.sub(r"^(86)", "", p))]
                phone = ",".join(set(phones))
        return phone

    def search(self):
        """��ں����������������"""
        sep_dir = re.split(r"/+|\\+", self.file_dir)
        directory = sep_dir[-2]
        file_name = sep_dir[-1]
        info = {"directory": directory, "file_name": file_name, "phone": "", "user_name": "", "email": ""}

        # ��������
        try:
            names = self.__search_name()
            info["user_name"] = ",".join(names)
        except Exception as e:
            print(e)

        # ����Email
        try:
            email = self.__search_email()
            info["email"] = email
        except Exception as e:
            print(e)

        # ���ҵ绰
        try:
            phone = self.__search_phone()
            info["phone"] = phone
        except Exception as e:
            print(e)
        return info

def find_files(file_dir):
    """���������ļ�"""
    file_paths = []
    for root, _, files in os.walk(file_dir):
        for file in files:
            path = os.path.join(root, file)
            rear = os.path.splitext(path)[1]
            if rear in [".doc", ".docx", ".pdf"]:
                file_paths.append(path)
    return file_paths

if __name__ == "__main__":
    FILE_DIR = r"data"
    OUT_DIR = r"resume-data.xlsx"
    args = sys.argv

    if len(args) > 1:
        FILE_DIR = args[1]
    if len(args) > 2:
        OUT_DIR = args[2]
        FILE_DIR = args[1]
    # �ļ����ڣ���׷�����
    cnt = 0
    while os.path.isfile(os.path.abspath(OUT_DIR)):
        OUT_DIR = os.path.splitext(OUT_DIR)[0] + "_" + str(cnt) + ".xlsx"
        cnt += 1
    writer = pd.ExcelWriter(OUT_DIR)
    for folder in os.listdir(FILE_DIR):
        file_dir = os.path.join(os.path.abspath(FILE_DIR), folder)
        paths = find_files(file_dir)
        print("Total {} file(s) in directory {}:".format(len(paths), folder))
        df = pd.DataFrame()
        for index, file_path in enumerate(paths):
            info = Extractor(file_dir=file_path).search()
            df = df._append(info, ignore_index=True)
            print(index, info["file_name"], info["email"], info["phone"], info["user_name"])
        df.to_excel(writer, folder)
    print("Save to file ", OUT_DIR)
    writer._save()
    print("All done.")

