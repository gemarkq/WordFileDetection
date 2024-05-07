import utils
from utils import WordStatus
import docx
from nltk.tokenize import word_tokenize, sent_tokenize
import os
import sys

class Matter:
    def __init__(self, path) -> None:
        self.path = path
        self.document = docx.Document(self.path)
        self.tables = self.document.tables
        self.paragraphs = []
        self.ExtractParagraphs()
    
    def ExtractParagraphs(self):
        for table in self.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self.paragraphs.append(para)

    def CheckClientName(self, client_name):
        first_name, last_name = client_name.split()
        for idx, para in enumerate(self.paragraphs):
            if not self.CheckName(para, first_name, last_name):
                print(f"[ERROR] paragraph: {idx}, {para.text}")
                return False
        return True

    def CheckName(self, para, client_first_name, client_last_name):
        # Attorney ShuKui GAO
        # Attorney ShuKui GAO’s
        # Attoreny GAO
        # Attorney GAO’s
        potentialc_client_name = [client_first_name + ' ' + client_last_name.upper(),
                                  client_first_name + ' ' + client_last_name.upper() + "'s",
                                  client_last_name.upper(),
                                  client_last_name.upper() + "'s"]
        words = word_tokenize(para.text)
        for i in range(1, len(words)):
            attorney_flag = utils.IsWordValid(words[i-1].lower(), "attorney")
            if attorney_flag == WordStatus.WORD_SPELL_CORRECT:
                first_name_flag = utils.IsWordValid(words[i], client_first_name)
                if first_name_flag == WordStatus.WORD_SPELL_CORRECT:
                    last_name_flag = utils.IsWordValid(words[i + 1], client_last_name)
                    if last_name_flag == WordStatus.WORD_SPELL_CORRECT:
                        pass
                    else:
                        print(f"[ERROR] Client last name spell wrong, {client_last_name}")
                        return False
                elif first_name_flag == WordStatus.WORD_SPELL_WRONG:
                    print(f"[ERROR] Client first name spell wrong, {words[i]}")
                    return False
            elif attorney_flag == WordStatus.WORD_SPELL_WRONG:
                print(f"[ERROR] Attorney spell wrong")
                return False
            else:
                pass
        return True

    def CheckPunctuation(self):
        for idx, para in enumerate(self.paragraphs):
            sentences = sent_tokenize(para.text)
            for sentence in sentences:
                words = word_tokenize(sentence)
                if not utils.IsValidParaPunctuation(words):
                    print(f"[ERROR] Use Chinese punctuation!!!, paragraph id: {idx}, sentence: {sentence}")
                    return False
        return True


if __name__ == "__main__":
    document_path = input(r"Please Input document file path: ")
    if not os.path.exists(document_path):
        print(f"{document_path} File Path does not exits. Please check!!!")
        sys.exit(0)
    client_name = input("Please Input Client Name (e.g. ShuKui GAO): ")
    client_name_check = input("Please ReInput Client Name: ")
    client_name = client_name.strip()
    client_name_check = client_name_check.strip()
    if client_name != client_name_check:
        print(f"Client Name input mismatch. Please re-input the client name")
        sys.exit(0)
    
    matter = Matter(document_path)
    if matter.CheckClientName(client_name):
        print(f"[Success] Client Name spell Correct!!! Congrats!")
    if matter.CheckPunctuation():
        print(f"[Success] No punctuation error!!! Congrats!")
    
    sys.exit(0)

