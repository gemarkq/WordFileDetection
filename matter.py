import utils
from utils import WordStatus
import docx
from nltk.tokenize import word_tokenize

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
                print(f"paragraph: {idx} client name wrong")
                print(para.text)

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
                        print(f"Client last name spell wrong, {client_last_name}")
                        return False
                elif first_name_flag == WordStatus.WORD_SPELL_WRONG:
                    print(f"Client first name spell wrong, {words[i]}")
                    return False
            elif attorney_flag == WordStatus.WORD_SPELL_WRONG:
                print(f"Attorney spell wrong")
                return False
            else:
                pass
        return True
                    
                


if __name__ == "__main__":
    matter = Matter(r'/Users/qiujiankun/Python_project/WordFileDetection/【Shirley】【完善版】高淑奎PM3-4.docx')
    matter.CheckClientName('ShuKui GAO')

