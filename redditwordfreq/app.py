'''
Importing all the dependencies for the script
'''
from bs4 import BeautifulSoup   # for parsing the html DOM element
import requests                 # to send request to the url
import xlsxwriter               # to write the output to excel file
from collections import Counter # to find the frequency of the words

headers = {'User-Agent':'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:54.0) Gecko/20100101 Firefox/54.0'}

workbook1 = xlsxwriter.Workbook('VotesTitle.xlsx')
worksheet1 = workbook1.add_worksheet()

workbook2 = xlsxwriter.Workbook('FrequencyList.xlsx')
worksheet2 = workbook2.add_worksheet()

votetitlecounter = 1
frequencycounter = 1

worksheet1.write(0,0,'No of Votes')
worksheet1.write(0,1,'Title')
worksheet2.write(0,0,'Word')
worksheet2.write(0,1,'Frequency')

'''
Input the subreddit title and send the request to the url and parsing with the help of beautifulsoup
'''

searchtitle = input('Enter the name of subreddit: ')
url = 'https://www.reddit.com/r/'+str(searchtitle)+'/hot'
myreq = requests.get(url, headers=headers)
mysoup = BeautifulSoup(myreq.text,'lxml')
mytoptenlist = []
alllist = ''

'''
Getting all the top 10 list and appending it to the mytoptenlist array
'''
for i in range(1,11):
    mytoptenlist.append(mysoup.find_all('div',{'data-rank':i}))

'''
Iterating over the mytoptenlist array and scraping the title and votes from them and writing in the 
excel file
'''
for a in range(len(mytoptenlist)):
    for i in mytoptenlist[a]:
        for j in i:
            for k in j:
                if 'score unvoted' in str(k):
                    print(k.text)
                    worksheet1.write(votetitlecounter,0,k.text)

                for l in k:
                    for m in l:
                        if 'title may-blank' in str(m):
                            print(m.text)
                            worksheet1.write(votetitlecounter,1,m.text)

                            '''Adding all the title in the alllist string to count the frequency of the 
                            word in the future
                            '''
                            alllist = str(alllist)+' '+m.text

    votetitlecounter += 1

'''
Calculating the frequency of each word in the string using the pre available library Counter.
Removing the '/' sign from the list and calculating the frequency of each word and order them
in the basis of decreasing frequency
'''
myfrequencyarray = Counter(alllist.replace("/","").split()).most_common()

'''
Writing the calculated frequency in the excel file Votestitle.xlsx
'''
for i in range(len(myfrequencyarray)):
    print(myfrequencyarray[i][0],myfrequencyarray[i][1])
    worksheet2.write(frequencycounter,0,myfrequencyarray[i][0])
    worksheet2.write(frequencycounter,1,myfrequencyarray[i][1])
    frequencycounter += 1
    # workbook.close()

'''
Closing the excel file when all the data are finished written.
'''
workbook1.close()
workbook2.close()
