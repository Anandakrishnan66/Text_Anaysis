import pandas as pd
import requests
import bs4 as bfs
import nltk
from nltk.tokenize import sent_tokenize
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.sentiment.vader import SentimentIntensityAnalyzer
import string
from textblob import TextBlob
import csv
import numpy as np
import openpyxl as op
import os
import re



input=pd.read_excel("./input.xlsx")

links=[url for url in input['URL']]

for index,row in input.iterrows():
    url=row['URL']
#     print(index)

    url_id=row["URL_ID"]
    
    links=links
    
    try:
        response=requests.get(url,headers={'User-Agent':"XY"})
        
    except:
        print("cant get the response of {}".format(url_id))
    
    try:
        soup=bfs.BeautifulSoup(response.content,'html.parser')
    except:
        print("cant get page of {}".format(url_id))
    
    try:
        title=soup.find('title').get_text()
    except:
        print("cant get title of {}".format(url_id))
        continue
    article=""
    
    try:
        for p in soup.find_all('p'):
            article+=p.get_text()
    except:
        print('can\'t get  article of {}'.format(url_id))
    
    
    file_name="./Text"+str(url_id)+".txt"
    
    with open(file_name,'w' ) as file:
        file.write(title + '\n'+ article)
        
    
text_dir="./Text"
stopwords_dir="./StopWords"
senti_dir="./MasterDictionary"

stop_words=set()
for files in os.listdir(stopwords_dir):
    with open(os.path.join(stopwords_dir,files),'r',encoding='ISO-8859-1') as f:
        stop_words.update(set(f.read().splitlines()))
# stop_words

text_files=[]
for text_file in os.listdir(text_dir):
    with open(os.path.join(text_dir,text_file),'r')as f:
        text=f.read()
        words=word_tokenize(text)
        filtered_text=[word for word in words if word.lower() not in stop_words]
        text_files.append(filtered_text)
        
pos=set()
neg=set()

for files in os.listdir(senti_dir):
    if files=='positive-words.txt':
        with open(os.path.join(senti_dir,files),'r',encoding='ISO-8859-1')as f:
            pos.update(f.read().splitlines())
    else:
        with open(os.path.join(senti_dir,files),'r',encoding='ISO-8859-1') as f:
            neg.update(f.read().splitlines())
            

positive_words=[]
negative_words=[]
positive_score=[]
negative_score=[]
polarity_score=[]
subjectivity_score=[]

for i in range(len(text_files)):
    positive_words.append([word for word in text_files[i] if word.lower() in pos])
    negative_words.append([word for word in text_files[i] if word.lower() in neg])
    positive_score.append(len(positive_words[i]))
    negative_score.append(len(negative_words[i]))
    
    polarity_score.append((positive_score[i]-negative_score[i])/((positive_score[i]+negative_score[i])+0.000001))
    
    subjectivity_score.append((positive_score[i]+negative_score[i]/(len(text_files[i]))+0.000001))
    

avg_sentence_length=[]
Percentage_of_complex_words=[]
Fog_index=[]
complex_word_count=[]
avg_syllable_word_count=[]

def measure(file):
    with open(os.path.join(text_dir,file),'r')as f:
            text=f.read()
            
            text=re.sub(r'[^\w\s.]','',text)
            
            sentences=text.split('.')
            
            num_sentences=len(sentences)
            
            words=[word for word in text.split() if word.lower() not in stop_words]
            num_words=len(words)
            
            
            complex_words=[]
            
            for word in words:
                vowels='aeiou'
                
                syllable_count_word=sum(1 for letter in word if letter .lower() in vowels)
                
                if syllable_count_word>2:
                    complex_words.append(word)
                    
            syllable_count=0
                
            syllable_words=[]
                
            for word in words:
                if word.endswith('es'):
                        word=word[:-2]
                        
                elif word.endswith('ed'):
                    word=word[:-2]
                vowels='aeiou'
                    
                syllable_count_word=sum(1 for letter in word if letter.lower() in vowels)
                    
                if syllable_count_word>=1:
                        syllable_words.append(word)
                        syllable_count+=syllable_count_word
                    
            avg_sentence_len=num_words/num_sentences
            avg_syllable_word_count=syllable_count/len(syllable_words)
            Percentage_complex_words=len(complex_words)/num_words
            Fog_index=0.4*(avg_sentence_len+Percentage_complex_words)
                
            return avg_sentence_len,Percentage_complex_words,Fog_index,len(complex_words),avg_syllable_word_count
            
        
            
for file in os.listdir(text_dir):
    x,y,z,a,b=measure(file)
    avg_sentence_length.append(x)
    Percentage_of_complex_words.append(y)
    Fog_index.append(z)
    complex_word_count.append(a)
    avg_syllable_word_count.append(b)
    
    
            
    
    
    
def cleaned_words(file):
    with open(os.path.join(text_dir,file),'r')as f:
        text=f.read()
        text=re.sub(r'[^\w\s]','',text)
        words=[word for word in text.split() if word.lower not in stop_words]
        length=sum(len(word) for word in words)
        average_word_length=length/len(words)
    return len(words),average_word_length

word_count=[]
average_word_length=[]

for file in os.listdir(text_dir):
    x,y=cleaned_words(file)
    word_count.append(x)
    average_word_length.append(y)
    
def count_personal_pronouns(file):
    with open(os.path.join(text_dir,file),'r')as f:
        text=f.read()
        personal_pronouns=["I","we","my","ours","us"]
        count=0
        for pronoun in personal_pronouns:
            count+=len(re.findall(r"\b"+pronoun+r"\b",text))
    return count

pp_count=[]

for file in os.listdir(text_dir):
    x=count_personal_pronouns(file)
    pp_count.append(x)
    

output_df=pd.read_excel('./Output Data Structure.xlsx')

variables = [positive_score,
            negative_score,
            polarity_score,
            subjectivity_score,
            avg_sentence_length,
            Percentage_of_complex_words,
            Fog_index,
            avg_sentence_length,
            complex_word_count,
            word_count,
            avg_syllable_word_count,
            pp_count,
            average_word_length]

for i ,var in enumerate(variables):
    output_df.iloc[:,i+2]=var
# output_df.to_csv('./Output.xlsx')
file_name="/Output.xlsx"
    
output_df.to_excel(file_name)


    
        
        
    
    
    




