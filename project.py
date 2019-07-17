import re
from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter

global keywordlist
keywordlist=[]
class project:
    def __init__(self,url):
        self.url=l        

    def script(self):
        ctr=0
        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        taglist=[]
        for tag in soup.find_all(re.compile("^script")):
            taglist.append(tag.name)
        print("number of scripts in webpage are:",len(taglist))
        
    def text(self):
        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        script=soup.find_all("script")
        for s in script:
            print(s)
        """ for i in soup:
            script = soup.find('script')
            print(script.text)
            break
        script = soup.find('script')
        for i in script:
            print(i)
        
        for s in soup(["script"]):
            print(s)"""
        
        
    def nwords(self):
        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        for s in soup(["script","style"]):
            s.extract()
        wordlist=soup.get_text().upper().split()
        print("number of words in webpage are:",len(wordlist))
        print("-"*30)
        
    def keyword(self):
        keywordlist=[]

        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        for s in soup(["script","style"]):
            s.extract()
        wordlist=soup.get_text().upper().split()
        metas=soup.find_all("meta")
        if len(metas)>0:
            for meta in metas:
                if "name" in meta.attrs and meta.attrs["name"].upper()=="KEYWORDS":
                    keywordlist=meta.attrs["content"].upper().split(",")
                    keywordlist.sort()
            print("Keywords in webpage ",self.url,"are:")
            for k in keywordlist:
                print(k)
        else:
            print("There is no keyword in webpage!!!")
        print("-"*30)
        
    def kfrequency(self):
        keyworddict={}
        keywordlist=[]
        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        for s in soup(["script","style"]):
            s.extract()
        wordlist=soup.get_text().upper().split()
        metas=soup.find_all("meta")
        if len(metas)>0:
            for meta in metas:
                if "name" in meta.attrs and meta.attrs["name"].upper()=="KEYWORDS":
                    keywordlist=meta.attrs["content"].upper().split(",")
                    keywordlist.sort()
            keyworddict={k:wordlist.count(k) for k in keywordlist}        
            for k,v in keyworddict.items():
                print("Keyword",k,"occurs",v,"times.")
        else:
            print("There is no keyword in webpage!!!")
            
            
        
    def xlsheet(self):
        keywordlist=[]
        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        for s in soup(["script","style"]):
            s.extract()
        wordlist=soup.get_text().upper().split()
        metas=soup.find_all("meta")
        for meta in metas:
            if "name" in meta.attrs and meta.attrs["name"].upper()=="KEYWORDS":
                keywordlist=meta.attrs["content"].upper().split(",")
                keywordlist.sort()
        keyworddict={k:wordlist.count(k) for k in keywordlist}
        workbook=xlsxwriter.Workbook("keysheet.xlsx")
        worksheet=workbook.add_worksheet()
        worksheet.write(0,0,"Keyword occurrences in "+self.url)
        worksheet.write(1,0,"Keyword")
        worksheet.write(1,1,"Occurrences")
        row=2
        for k,v in keyworddict.items():
            worksheet.write(row,0,k)
            worksheet.write(row,1,v)
            row+=1
        workbook.close()
        print("Created Excel sheet :keysheet.xlsx")
        
    def chart(self):
        keywordlist=[]
        fh=urlopen(self.url)
        html=fh.read()
        fh.close()
        soup=BeautifulSoup(html,"html.parser")
        for s in soup(["script","style"]):
            s.extract()
        wordlist=soup.get_text().upper().split()
        metas=soup.find_all("meta")
        for meta in metas:
            if "name" in meta.attrs and meta.attrs["name"].upper()=="KEYWORDS":
                keywordlist=meta.attrs["content"].upper().split(",")
                keywordlist.sort()
        keyworddict={k:wordlist.count(k) for k in keywordlist}
        workbook=xlsxwriter.Workbook("keychart.xlsx")
        worksheet=workbook.add_worksheet()
        worksheet.write(0,0,"Keyword occurrences in "+self.url)
        worksheet.write(1,0,"Keyword")
        worksheet.write(1,1,"Occurrences")
        row=2
        for k,v in keyworddict.items():
            worksheet.write(row,0,k)
            worksheet.write(row,1,v)
            row+=1
        chart=workbook.add_chart({"type":"column"})
        chart.add_series({"name":"Occurrences","values":"=Sheet1!$B$3:$B$"+str(row+1),"categories":"=Sheet1!$A$3:$A$"+str(row+1)})
        worksheet.insert_chart(0,3,chart)
        workbook.close()
        print("Created Excel sheet with chart keychart.xlsx")
        print("-"*30)
        
        
keywordlist=[]
l=input("Enter URL:")
p=project(l)
f=1
while f==1:
    print("_"*50)
    
    print("""1.Find total number of scripts.
2. Display a list of the script text.
3. Find total number of words excluding scripts and styles.
4. Display all keywords from meta tag.
5. Display all keywords along with their number of occurrences in the page(using dictionary).
6.Create an excel sheet.
7. Create column chart.
8.Exit.""")
    print("-"*30)
    choice=int(input("Enter your choice:"))

    if choice==1:
        p.script()
    elif choice==2:
        p.text()
    elif choice==3:
        p.nwords()
    elif choice==4:
        p.keyword()
    elif choice==5:
        p.kfrequency()
    elif choice==6:
        p.xlsheet()
    elif choice==7:
        p.chart()
    elif choice==8:
        break
    else:
        print("Invalid Choice.")
    
    
