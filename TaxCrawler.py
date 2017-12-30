import requests
from bs4 import BeautifulSoup
import re
import collections
import xlwt
import time
import yagmail
import glob, os


#puts link to all account into array
def getAccountLinks(s):

        site=s
        with open(site, 'r') as myfile:
            plain_text= myfile.read().replace('\n', '')
        soup = BeautifulSoup(plain_text, 'html.parser')
        temp=list()
        for link in soup.findAll():  # {'class': 's-result-item celwidget '}
            href = link.get("href")
            if href != None and "showdetail2" in href:
                temp.append(href)
                #print(href)
        return temp
#returns a list of list of account data
def getAccountData(url):
    # later for webpage
    source_code = requests.get(url, allow_redirects=False)
    plain_text = source_code.text.encode('ascii', 'replace')
    soup = BeautifulSoup(plain_text, 'html.parser')
    i=0
    all=[]

    for link in soup.findAll("h3"):  # {'class': 's-result-item celwidget '}
        all+=re.split(r'\t|\xa0|  |<[^>]*>',str( link))

    all=[x for x in all if x and len(x)>2]
    accdata=[]
    accdata.append(all[0:all.index("Address:")])
    accdata.append(all[all.index("Address:"):all.index("Property Site Address:")])
    accdata.append(all[all.index("Property Site Address:"):all.index("Legal Description:")])
    accdata.append(all[all.index("Legal Description:"):all.index("Current Tax Levy: ")])
    accdata.append(all[all.index("Current Tax Levy: "):all.index("Current Amount Due: ")])
    accdata.append(all[all.index("Current Amount Due: "):all.index("Prior Year Amount Due: ")])
    accdata.append(all[all.index("Prior Year Amount Due: "):all.index("Total Amount Due: ")])
    accdata.append(all[all.index("Total Amount Due: "):all.index("Total Amount Due: ")+2])
    accdata.append(all[all.index("Market Value:"):all.index("Land Value:")])
    accdata.append(all[all.index("Land Value:"):all.index("Improvement Value:")])
    accdata.append(all[all.index("Improvement Value:"):all.index("Capped Value:")])
    accdata.append(all[all.index("Capped Value:"):all.index("Agricultural Value:")])
    accdata.append(all[all.index("Agricultural Value:"):all.index("Exemptions:")])
    accdata.append(all[all.index("Exemptions:"):all.index("Exemptions:")+2])

    return accdata


#this part involves writing data
wb = xlwt.Workbook()
#gets page with links
def writeAccountData(file_name):

    sheet1 = wb.add_sheet(file_name)

    row=0
    for link in getAccountLinks(file_name):
        curracc = getAccountData(link)
        if (curracc[7][1] == "$0.00"):
            continue
        sheet1.write(row ,0," ".join(curracc[0][1:]))
        sheet1.write(row, 1," ".join(curracc[1][1:]))
        sheet1.write(row, 2," ".join(curracc[2][1:]))
        sheet1.write(row, 3," ".join(curracc[3][1:]))
        sheet1.write(row, 4," ".join(curracc[4][1:]))
        sheet1.write(row, 5," ".join(curracc[5][1:]))
        sheet1.write(row, 6," ".join(curracc[6][1:]))
        sheet1.write(row, 7," ".join(curracc[7][1:]))
        sheet1.write(row, 8," ".join(curracc[8][1:]))
        sheet1.write(row, 9," ".join(curracc[9][1:]))
        sheet1.write(row, 10," ".join(curracc[10][1:]))

        sheet1.write(row, 11," ".join(curracc[11][1:]))
        sheet1.write(row, 12," ".join(curracc[12][1:]))
        sheet1.write(row, 13," ".join(curracc[13][1:]))

        row += 1
#gets all file names from directory
def printFromDirectory():

    # os.chdir("./StreetHTML")
    for file in glob.glob("*.htm"):
        try:
            writeAccountData(file)
        except :
            print(file)
    wb.save('All Streets.xls')

printFromDirectory()






#later helpful functions

#writes each account into different spreadsheet??helpful later
def writeAccountDataSheets(url):
    # text_file = open("Output.txt", "w")
    # text_file.write
    wb = xlwt.Workbook()


    row=0
    for link in getAccountLinks(url):

        curracc = getAccountData(link)
        if (curracc[7][1:] == "$0.00"):
            break
        sheet1 = wb.add_sheet(str(curracc[0][1]))
        sheet1.write(row ,0," ".join(curracc[0][1:]))
        sheet1.write(row, 1," ".join(curracc[1][1:]))
        sheet1.write(row, 2," ".join(curracc[2][1:]))
        sheet1.write(row, 3," ".join(curracc[3][1:]))
        sheet1.write(row, 4," ".join(curracc[4][1:]))
        sheet1.write(row, 5," ".join(curracc[5][1:]))
        sheet1.write(row, 6," ".join(curracc[6][1:]))
        sheet1.write(row, 7," ".join(curracc[7][1:]))
        sheet1.write(row, 8," ".join(curracc[8][1:]))
        sheet1.write(row, 9," ".join(curracc[9][1:]))
        sheet1.write(row, 10," ".join(curracc[10][1:]))

        sheet1.write(row, 11," ".join(curracc[11][1:]))
        sheet1.write(row, 12," ".join(curracc[12][1:]))
        sheet1.write(row, 13," ".join(curracc[13][1:]))


    wb.save('results.xls')
#check if account exists
def checkIfAcc(url):
    source_code = requests.get(url, allow_redirects=False)
    plain_text = source_code.text.encode('ascii', 'replace')
    soup = BeautifulSoup(plain_text, 'html.parser')
    i = 0
    all = []

    for link in soup.findAll("h6"):  # {'class': 's-result-item celwidget '}
        all += re.split(r'\t|\xa0| |<[^>]*>', str(link))

    return ("error" not in all)
#tries to get every combination
import itertools
def getAllAccountData():
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet('A Test Sheet')
    row = 0
    for combination in itertools.product(range(10), repeat=6):
        accountnum=(''.join(map(str, combination)))
        print(accountnum)
        if(checkIfAcc("http://www.dallasact.com/act_webdev/dallas/showdetail2.jsp?can=00000"+accountnum+"000000&ownerno=0")):
            print("we found some")
            curracc = getAccountData("http://www.dallasact.com/act_webdev/dallas/showdetail2.jsp?can=00000"+accountnum+"000000&ownerno=0")
            sheet1.write(row, 0, " ".join(curracc[0][1:]))
            sheet1.write(row, 1, " ".join(curracc[1][1:]))
            sheet1.write(row, 2, " ".join(curracc[2][1:]))
            sheet1.write(row, 3, " ".join(curracc[3][1:]))
            sheet1.write(row, 4, " ".join(curracc[4][1:]))
            sheet1.write(row, 5, " ".join(curracc[5][1:]))
            sheet1.write(row, 6, " ".join(curracc[6][1:]))
            sheet1.write(row, 7, " ".join(curracc[7][1:]))
            sheet1.write(row, 8, " ".join(curracc[8][1:]))
            sheet1.write(row, 9, " ".join(curracc[9][1:]))
            sheet1.write(row, 10, " ".join(curracc[10][1:]))

            sheet1.write(row, 11, " ".join(curracc[11][1:]))
            sheet1.write(row, 12, " ".join(curracc[12][1:]))
            sheet1.write(row, 13, " ".join(curracc[13][1:]))
            row+=1
    wb.save('results.xls')






# later for webpage
# url = 'https://www.--------------.com'
# source_code = requests.get(url, allow_redirects=False)
# just get the code, no headers or anything
# plain_text = source_code.text.encode('ascii', 'replace')
# writeAccountData('Douglas Ave.html')
# wb.save('All Streets.xls')
#sandpaper ln
#getAllAccountData()
# print(checkIfData("http://www.dallasact.com/act_webdev/dallas/showdetail2.jsp?can=99130425120000000&ownerno=0"))
#writeAccountData('Dallas County Web Site.html')
#
# def bbc():
#     url = 'http://www.bbc.com/news'
#     source_code = requests.get(url, allow_redirects=False)
#     # just get the code, no headers or anything
#     plain_text = source_code.text.encode('ascii', 'replace')
#     # BeautifulSoup objects can be sorted through easy
#     soup = BeautifulSoup(plain_text, 'html.parser')
#     temp = list();
#     # print(soup.get_text())
#     for link in soup.findAll("a",{'class': 'gs-c-promo-heading nw-o-link-split__anchor gs-o-faux-block-link__overlay-link gel-pica-bold'}):  # {'class': 's-result-item celwidget '}
#         link = str(link)
#
#         if "http" in link:
#             temp.append(link[link.find("href") + 5:])
#         else:
#             temp.append(('http://www.bbc.com/news' + link[link.find("href") + 6:]))
#
#     return temp
# #gets all new headlines into allnews ################################################333333333333333
# def allNews():
#     allNews=econ()+nyt()+wsj()+bbc()
#     allNews=[x.lower() for x in allNews]
#     print(len(allNews))
#     return allNews
# #divides headlines into words
# def words():
#     keywords=list()
#     headlines=allNews()
#     for i in range(len(headlines)):
#         keywords += re.split(r'[<>/-]+| ', headlines[i])
#     return keywords
# #gets the keywords from the words and orders them
# def keywords():
#     wordlist=words()
#     counter = collections.Counter(wordlist)
#     for e in ['make','science','media','07','split__text">the','32','12,''vh@xs','inline"><span','span><h3','svg><','mr','viewbox="0','blogs.wsj.com','top','16','span><','aria','hidden="true"','special','report', 'story','indicator','politics','opinion','The','all', 'just', 'being', 'over', 'both', 'through', 'yourselves', 'its', 'before', 'herself', 'had', 'should', 'to', 'only', 'under', 'ours', 'has', 'do', 'them', 'his', 'very', 'they', 'not', 'during', 'now', 'him', 'nor', 'did', 'this', 'she', 'each', 'further', 'where', 'few', 'because', 'doing', 'some', 'are', 'our', 'ourselves', 'out', 'what', 'for', 'while', 'does', 'above', 'between', 't', 'be', 'we', 'who', 'were', 'here', 'hers', 'by', 'on', 'about', 'of', 'against', 's', 'or', 'own', 'into', 'yourself', 'down', 'your', 'from', 'her', 'their', 'there', 'been', 'whom', 'too', 'themselves', 'was', 'until', 'more', 'himself', 'that', 'but', 'don', 'with', 'than', 'those', 'he', 'me', 'myself', 'these', 'up', 'will', 'below', 'can', 'theirs', 'my', 'and', 'then', 'is', 'am', 'it', 'an', 'as', 'itself', 'at', 'have', 'in', 'any', 'if', 'again', 'no', 'when', 'same', 'how', 'other', 'which', 'you', 'after', 'most', 'such', 'why', 'a', 'off', 'i', 'yours', 'so', 'the', 'having', 'once']:
#         counter.pop(e,"yes")
#     orderedwords=counter.most_common()
#     # removes the 48 useless
#     for i in range(48):
#         orderedwords.pop(0)
#     return orderedwords
# #search for news headline
# def keywordSearch():
#     keysearch = input('Enter keyword:')
#     keysearch=keysearch.lower()
#     print('\n\n\n\n\n\n')
#     headlines=allNews()
#     for line in headlines:
#         if(keysearch in line):
#             print(line)
#for text file
    # for link in getAccountLinks(url):
    #     curracc=getAccountData(link)
    #     text_file.write("%s - %s - %s - %s - %s - %s - %s - %s - %s - %s - %s - %s - %s - %s" %
    #           (
    #
    #             " ".join(curracc[0][1:])," ".join(curracc[1][1:]), " ".join(curracc[2][1:]),
    #            " ".join(curracc[3][1:])," ".join(curracc[4][1:])," ".join(curracc[5][1:]),
    #             " ".join(curracc[6][1:])," ".join(curracc[7][1:])," ".join(curracc[8][1:]), " ".join(curracc[9][1:]),
    #            " ".join(curracc[10][1:])," ".join(curracc[11][1:])," ".join(curracc[12][1:]),
    #             " ".join(curracc[13][1:]
    #
    #
    #
    #           )))
    # text_file.close()