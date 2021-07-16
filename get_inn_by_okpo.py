import win32com
import requests
from bs4 import BeautifulSoup

class OmniWebscrapper():
    # com spec
    _public_methods_ = ['get_inn_by_okpo',] # методы объекта
    _public_attrs_ = ['version',] # атрибуты объекта
    _readonly_attr_ = []
    _reg_clsid_ = '{5bbc4b66-da07-48bb-beb1-f457a8db0e87}' # uuid объекта
    _reg_progid_= 'OmniWebscrapper' # id объекта
    _reg_desc_  = 'get organization data from web-sites (cuurently list-org.com)' # описание объекта
    
    def __init__(self):
        self.version = '0.0.1'
        # ...

    def get_inn_by_okpo(self, okpo):
        r = requests.get(f"https://www.list-org.com/search?type=okpo&val={okpo}", 
                 headers={'User-agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:61.0) Gecko/20100101 Firefox/61.0'})
        #маскируемся под браузер, чтобы сайт нас не прогнал
        c = r.content

        soup = BeautifulSoup(c, "html.parser")
        spans = soup.find_all("span")
        for cur_span in spans:    
            for cur_i in cur_span.find_all("i"):
                cur_text = cur_i.next_sibling
                if "инн" in cur_i.text:
                    return cur_text.replace(":", "").strip()


def main():
    import win32com.server.register
    win32com.server.register.UseCommandLine(OmniWebscrapper)
    print('registered')

if __name__ == '__main__':
    main()

    






# Процедура Кнопка1Нажатие(Элемент)
#     //Создадим объект
#    Пример =  Новый COMОбъект("ExampleCOM");
#    
#    Сообщить(Пример.hello("Daniel"));
# КонецПроцедуры