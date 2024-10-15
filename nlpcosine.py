import docx2txt
import warnings
import os
import spacy
import pandas as pd
import en_core_web_sm
import webbrowser
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from spacy.matcher import PhraseMatcher
from prettytable import PrettyTable
from collections import OrderedDict

warnings.filterwarnings('ignore')

nlp=spacy.load("en_core_web_sm")
mypath= "D:\lawda\Resume-Scanner-for-Job-Description--using-Cosine-Similarity\Resumes"
os.chdir(mypath)
cosarray=[]
phrarray=[]
finarray=[]
filename=[]
print("Processing...")

for file in os.listdir():
    finpath = file
    filename.append(finpath)
    resume = docx2txt.process(finpath)
    job = docx2txt.process('D:\lawda\Resume-Scanner-for-Job-Description--using-Cosine-Similarity\python-job-description.docx')
    text = [resume,job]
    cv = CountVectorizer()
    count_matrix = cv.fit_transform(text)
    perc = cosine_similarity(count_matrix)[0][1]
    perc = round(perc*100,2)
    cosarray.append(perc)

matcher = PhraseMatcher(nlp.vocab)
for file in os.listdir():
    finpath = file
    resume = docx2txt.process(finpath)
    resume = resume.split(" ")
    patterns = [nlp(test) for test in resume]
    matcher.add("RESUME_PATTERN", patterns)
    keyword = docx2txt.process('D:\lawda\Resume-Scanner-for-Job-Description--using-Cosine-Similarity\kword.docx')
    doc = nlp(keyword)
    matches = matcher(doc)
    count=0
    for match_id, start, end in matches:
        span = doc[start:end]
        count=count+1
    phrarray.append((count/len(keyword.split(" ")))*100)
        
for i in range(len(cosarray)):
    finarray.append(((cosarray[i]*7)+(phrarray[i])*3)/10)
   
dictionary=dict()
for i in range(len(filename)):
    dictionary[filename[i]]=finarray[i]

sorted_values = sorted(dictionary.values())
sorted_dict = {}
for i in sorted_values:
    for k in dictionary.keys():
        if(dictionary[k]==i):
            sorted_dict[k]=dictionary[k]
            break

final_dict = dict(OrderedDict(reversed(list(sorted_dict.items()))))
final_list = list(final_dict.items())
x = PrettyTable()
x.field_names=["Resume", "Score"]
for i in range(len(final_list)):
    file = str(final_list[i][0])
    resume1 = file.replace('.docx', '')
    x.add_row([resume1, final_list[i][1]])
print(x)

f = open('nlpresult.html', 'w')
html_temp= """<html>
<head>
<title> NLP Project </title>
</head>
<body>
"""
f.write(html_temp)
html=x.get_html_string()
f.write(html)
html_temp = """</body></html>"""
f.write(html_temp)
f.close()
webbrowser.open('nlpresult.html')