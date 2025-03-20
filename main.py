from autobib import get_bibtext, entries_to_str
import bibtexparser
from bibtexparser.bparser import BibTexParser
import numpy as np
import pandas as pd
import math
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import requests
import xmltodict
from bs4 import BeautifulSoup as BS

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

criteria_cat1 = ['embodied conversational agent', 'ECA', 'conversation agent','small talk', 'speech-based', 'voice-based','verbal','chatbot','communication','dialog','dialogue']

criteria_cat2 = ['virtual human','digital human','virtual avatar','virtual character','virtual agent','virtual assistant','agent','avatar','character','embodied agent','assistant', 'hmanoid']

criteria_cat3 = ['develop','implement','build','create','design','system','application','development','implementation']
criteria_cat4 = ['user study','experiment','interaction','within group','between group','participant','testing','evaluate','assess','perception','human-computer interaction','human-agent interaction','virtual human interaction', 'hci','hai','human-ai interaction','experience', 'affordance','usability']

series = ['VRST', 'CHI', 'ICVR', 'XR', 'ISMAR', 'SIGGRAPH', 'VR', 'EGVE']

xr_terms = ['extended reality', 'virtual reality', 'augmented reality', 'mixed reality', 'xr','vr','ar','mr','virtual environment', 'virtual world', 'digital replica', 'digital twin', 'head-mounted device','hmd','immersive','360']
llm_terms = ['llm','large language model','bert', 'chatgpt','lambda','claude','cohere','ernie','falcon 40b','gemini','gemma','gpt','llama','mistral','orca','palm','stablelm','vicuna 33b','phi-1']
reviewpaper_terms= ['review paper','systematic review','a review of','review']

def get_authors_name(authors):
    author_list = ""
    for author in authors:
        if 'family' in author.keys():
            author_list += f"{author['family']}, {author['given'][0]}., "
    return author_list

def open_all_doi(file):
    data = pd.read_csv(file, encoding_errors='ignore')
    data = data[data['abstract'].isna()]

    dois = data['doi'].array

    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)

    options = Options()
    options.add_experimental_option("detach", True)
    s = Service()
    webBrowser = webdriver.Chrome(service=s, options=options)

    for i in range(0, 35):  
        webBrowser.get(f'https://doi.org/{dois[i]}')
        webBrowser.execute_script("window.open('');")
        time.sleep(3)
        webBrowser.switch_to.window(webBrowser.window_handles[(i+1)])

def convert_entry(value):

    if(isinstance(value, float) ):
        return str(int(value))
    
    return str(value)

def from_csv_bib_entries(file, output):
    df = pd.read_csv(file, encoding_errors='ignore')
    data = df.to_numpy()
    print(data)

    columns = df.columns.to_list()

    entries = []

    entry_type = [i for i in range(len(columns)) if columns[i] == 'ENTRYTYPE'][0]
    entry_id = [i for i in range(len(columns)) if columns[i] == 'ID'][0]

    for entry in data:
        bib = "@" + entry[entry_type] + "{" + f"{entry[entry_id]},\n"
        for i in range(1, len(columns) - 1):
            if(i != entry_type and i != entry_id and str(entry[i]) != 'nan' and entry[i] != float('NaN')):
                bib += f"{columns[i]} = " + "{" + convert_entry(entry[i]) + "},\n"

        last = len(columns) - 1
        if(str(entry[last]) != 'nan' and entry[last] != float('nan')):
            bib += f"{columns[last]} =" + "{"+ convert_entry(entry[last]) + "}\n}"
        else:
            bib = bib[:-2] + "\n}\n"
        
        entries.append(bib)

    with open(output,'w', encoding="utf-8") as tfile:
	    tfile.write('\n'.join(entries))

def compare_decision():
    df_PA = pd.read_excel("screening-articles-PA.xlsx", index_col=0) 
    df_CM = pd.read_excel("screening-articles-CM.xlsx", index_col=0) 

    df_PA['decision_PA'] = df_PA['Decision']
    df_PA['decision_CM'] = df_CM['Decision']

    conditions = [(df_PA['decision_CM'] == 'YES') & (df_PA['decision_PA'] == 'YES'), (df_PA['decision_CM'] == 'YES') | (df_PA['decision_PA'] == 'YES'), (df_PA['decision_CM'] == 'NO') & (df_PA['decision_PA'] == 'NO')]

    choices = ['YES', 'CHECK', 'NO']

    df_PA['final_decision'] = np.select(conditions, choices, default='NO') 

    df_PA.to_excel("screening-articles-FINAL.xlsx")

def open_urls(file):
    data = pd.read_excel(file)
    
    urls = data[data.after_discussion == 'YES']['url'].array

    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)

    options = Options()
    options.add_experimental_option("detach", True)
    options.add_experimental_option("detach", True)
    s = Service()
    webBrowser = webdriver.Chrome(service=s, options=options)

    for i in range(0, len(urls)):  
        webBrowser.get(urls[i])
        webBrowser.execute_script("window.open('');")
        time.sleep(3)
        webBrowser.switch_to.window(webBrowser.window_handles[(i+1)])

def filter_selected_papers():
    df_final = pd.read_excel("screening-articles-FINAL_discussed.xlsx")
    df_filter = pd.read_excel("screening-papers-extracted-info.xlsx")

    print(df_filter.columns)

    selected_papers = df_filter[df_filter["decision"] == "YES"]["doi"].tolist()

    df_final = df_final[df_final["doi"].isin(selected_papers)]
    df_final.to_excel("included-articles.xlsx")
    
def combine_files(shared_name, n, ext):
    data = ""
    for i in range(n):
        with open(f'{shared_name}_{(i+1)}.{ext}', encoding="utf8") as file_content:
            data += file_content.read()
    with open(f"combined-file.{ext}", "w", encoding="utf8") as outfile:
        outfile.write(data)

def get_bib_list_from_csv(file):

    bib_entries = ''
    data = pd.read_csv(file)


    #for value in data['DOI']:
    for value in data['Item DOI']:
        if (isinstance(value, str)):
            bib_entries += get_bibtext(value) + '\n'

    with open(f"output.bib", "w", encoding="utf8") as outfile:
        outfile.write(bib_entries)

def calculated_pages(bib):

    pages = 0

    if ('pages' in bib.keys()):

        entry = bib['pages']

        if('--' in entry):
            a,b = entry.split('--')
            if a.isnumeric():
                pages = int(b) - int(a)
        else:
            if entry.isnumeric():
                pages = int(entry)
    return pages

def words_in_text(words, text):
    return any([x in text for x in words])

def calculated_matching_score(bib):

    score = 0
    has_xr = False
    has_dev = False
    has_study = False
    has_llm = False
    is_review = False


    for field in bib.keys():
        cat1 = len([word for word in criteria_cat1 if(word in bib[field].lower())])
        cat2 = len([word for word in criteria_cat2 if(word in bib[field].lower())])
        cat3 = len([word for word in criteria_cat3 if(word in bib[field].lower())])
        cat4 = len([word for word in criteria_cat4 if(word in bib[field].lower())])


        #check if has XR
        if(words_in_text(xr_terms, bib[field].lower())):
            has_xr = True
        
        #check if has its implementaion
        if(words_in_text(criteria_cat3, bib[field].lower())):
            has_dev = True

        #check if has user study
        if(words_in_text(criteria_cat4, bib[field].lower())):
            has_study = True

        #check if has LLM involved
        if(words_in_text(llm_terms, bib[field].lower())):
            has_llm = True

        #check if it's a review paper
        if(words_in_text(reviewpaper_terms, bib[field].lower())):
            is_review = True

        #make journal higher priority
        if(field == 'journal'):
            score += 10

        if(field == 'series'):
            if any([x in bib[field].lower() for x in series]):
                score += 10

        #score += (cat1 + cat2)
        score += (cat1 + cat2 + cat3 + cat4)

    #exclude if not XR
    if (has_xr):
        score += 5
    else:
        score = 0
    
    
    #exclude if it's a review paper --> or select for review
    #if (!is_review):
    #    score = 0

    bib['XR'] = str(has_xr)
    bib['Dev'] = str(has_dev)
    bib['Study'] = str(has_study)
    bib['LLM'] = str(has_llm)
    bib['Review'] = str(is_review)
    bib['score'] = score


def bib_remove_no_author(file, id):
    with open(file, encoding="utf8") as file_content:
        data = file_content.readlines()

        print(data)

        entries = []
        entry = ''

        for line in data:
            entry += line
            if ('}\n' == line):
                entries.append(entry)
                entry = ''

        parsedbib = []
        
        for bib in entries:
            print("**"+bib)
            parser = BibTexParser()
            parser.ignore_nonstandard_types = False
            bibdb = bibtexparser.loads(bib, parser)
            #print(bibdb.entries)
            entry, = bibdb.entries

            #exclude if no author entry 
            #will also exclude proceedings to remove duplicates -> no authors in particular
            if('author' in entry.keys() and entry['author']!={}):
                parsedbib.append(entry)
        
        print(len(entries))
        print(len(parsedbib))

        with open(f"combined_remove_no_author.bib", "w", encoding="utf8") as outfile:
            outfile.write(entries_to_str(parsedbib))

def bib_remove_review_paper(file, id):
    return 0

def get_csv_from_bib(file, id):
    with open(file, encoding="utf8") as file_content:
            data = file_content.readlines()

            print(data)

            entries = []
            entry = ''

            for line in data:
                entry += line
                if ('}\n' == line):
                    entries.append(entry)
                    entry = ''

            print(len(entries))
            parsedbib = []

            for bib in entries:
                

                parser = BibTexParser()
                parser.ignore_nonstandard_types = False
                bibdb = bibtexparser.loads(bib, parser)
                print(bibdb.entries)
                entry, = bibdb.entries

                if('title' in entry.keys()):
                    entry['title'] = entry['title'].replace('{', '').replace('}', '')
                
                parsedbib.append(entry)

            file_csv = pd.DataFrame(parsedbib)
            file_csv.to_csv(f'screening-articles-{id}.csv')

        
        

def read_bib_entries(file, id):
    with open(file, encoding="utf8") as file_content:
        data = file_content.readlines()

        print(data)

        entries = []
        entry = ''

        for line in data:
            entry += line
            if ('}\n' == line):
                entries.append(entry)
                entry = ''

        print(len(entries))
        parsedbib = []
        reviewbib = []

        for bib in entries:
            print("**"+bib)

            parser = BibTexParser()
            parser.ignore_nonstandard_types = False
            bibdb = bibtexparser.loads(bib, parser)
            print(bibdb.entries)
            entry, = bibdb.entries

            if('title' in entry.keys()):
                entry['title'] = entry['title'].replace('{', '').replace('}', '')

            calculated_matching_score(entry)
            
            #exclude if pages < 5
            entry['pages_num'] = calculated_pages(entry)
            if(entry['pages_num'] > 0):
                if(entry['pages_num'] < 5):
                    entry['score'] = 0

            entry['pages_num'] = str(entry['pages_num'])

            #exclude if no author entry 
            if('author' not in entry.keys()):
                entry['score'] = 0
            else:
                entry['author'] = entry['author'].replace("{", "").replace("}", "")
            
            #exclude if no doi
            if('doi' not in entry.keys()):
                entry['score'] = 0
            
            #exclude if no scopus id
            if('scopus_id' not in entry.keys()):
                entry['score'] = 0

            #exclude if publication date earlier than 2014 
            if('year' in entry.keys()):
                if(int(entry['year']) < 2014):
                    entry['score'] = 0
            
            # #only include if ACR >1.5
            # if(entry['ACR'] < 1.5):
            #     entry['score'] = 0
            
            #only include if score > 0 and not review papers
            if(entry['score'] > 0 and entry['Review']=='False'):
                parsedbib.append(entry)

            #save to another csv for review papers
            if(entry['score'] > 0 and entry['Review']=='True'):
                reviewbib.append(entry)

            entry['score'] = str(entry['score'])

        parsedbib = sorted(parsedbib, key=lambda x: int(x['score']), reverse=True)
        reviewbib = sorted(reviewbib, key=lambda x: int(x['score']), reverse=True)


        print(len(parsedbib))
        print(len(reviewbib))

        with open(f"screening-articles-{id}.bib", "w", encoding="utf8") as outfile:
            outfile.write(entries_to_str(parsedbib))

        file_csv = pd.DataFrame(parsedbib)
        file_csv.to_csv(f'screening-articles-{id}.csv')
        #print(file_csv)
        # file_csv = file_csv[['title', 'year', 'keywords', 'abstract', 'journal', 'author', 'VR', 'video', 'DVR', 'school', 'score', 'publisher', 'series', 'pages', 'booktitle','month', 'doi', 
        #                       'isbn', 'ENTRYTYPE', 'ID', 'pages_num', 'issn', 'address', 'editor', 'url', 'number']]

        file2_csv = pd.DataFrame(reviewbib)
        file2_csv.to_csv(f'screening-review-papers-{id}.csv')


        
def get_doi_from_title(file):
    data = pd.read_csv(file, encoding_errors='ignore')
    title_entries = data['title'].values
    doi_entries = data['doi'].values
    doi_entries = [f"{str(doi).replace('https://doi.org/', '')}" for doi in doi_entries]
    count = 0
    for i in range(len(doi_entries)):
        if str(doi_entries[i]) == 'nan':
            count += 1
            url = f"https://api.crossref.org/works?query.bibliographic={title_entries[i]}"
            response = requests.get(url)
            if response.status_code == 200 :
                content = response.json()
                if content['message']['items']:
                    doi_entries[i] = content['message']['items'][0]['DOI']
                    print(doi_entries[i])
            else:
                print("Error API request")
    print(count)
    data['doi'] = doi_entries
    data.to_csv(f'{file}', index=False)

def cite_by_from_scopus(file):

    data = pd.read_csv(file, encoding_errors='ignore')

    doi_entries = data['doi'].values
    #year_entries = data['year'].values
    
    cite_by = ['' for i in range(len(doi_entries))]

    #cite_by = data['cite_by'].values

    for i in range(len(doi_entries)):

        url = f'http://api.elsevier.com/content/search/scopus?query=DOI({doi_entries[i]})'
        #url = f'https://api.elsevier.com/content/abstract/scopus_id/85042030188'

        # Making a get request 
        response = requests.get(url, headers={"Accept": "application/xml", "X-ELS-APIKey": "fafe1117886418fdc64ee21c0347792a"})  
        # print json content
        print(url)
        if(response.status_code == 200):
            #print(response.content)
            soup = BS(response.content, features="xml")
            
            if soup.find('citedby-count') is not None:
                cite_by[i] = soup.find('citedby-count').text
            else:
                cite_by[i] = 0
        else:
            cite_by[i] = 0


    data['cite_by'] = cite_by
    #data['ACR'] = [(cite_by[i]/(2024 - year_entries[i])) for i in range(len(year_entries))]

    data.to_csv(f'{file}', index=False)

def calculate_ACR(file, year):
    data = pd.read_csv(file, encoding_errors='ignore')
    year_entries = data['year'].values
    cite_by_entries = data['cite_by'].values
    ACR_entries = ['' for i in range(len(year_entries))]
    for i in range(len(year_entries)):
        ACR_entries[i] = (float)(cite_by_entries[i] / (year-year_entries[i]))

    data['ACR'] = ACR_entries
    data.to_csv(f'{file}', index=False)
        
        

def scopus_id(file):
    data = pd.read_csv(file, encoding_errors='ignore')
    doi_entries = data['doi'].values
    doi_entries = [f"{str(doi).replace('https://doi.org/', '')}" for doi in doi_entries]
    scopus_ids = ["" for i in  range(len(doi_entries))]
    #scopus_ids = data['scopus_id'].values
    count = 0
    #for i in range(len(doi_entries)):
    for i in range(499,500):
        count += 1
        url = f'http://api.elsevier.com/content/search/scopus?query=DOI({doi_entries[i]})'
        response = requests.get(url, headers={"Accept": "application/xml", "X-ELS-APIKey": "b5131d0ad91b62931d30b30f2e2252df"})
        print(response.content)
        if(response.status_code == 200):
            soup = BS(response.content, features="xml")
            if soup.find('dc:identifier') is not None:
                scopus_ids[i] = soup.find('dc:identifier').text.replace("SCOPUS_ID:", "")
                print(scopus_ids[i])
    #print(count)        
    #data["scopus_id"] = scopus_ids
    #data.to_csv(f'{file}', index=False)

def abstract_from_scopus(file):
    data = pd.read_csv(file, encoding_errors='ignore')
    data = data.fillna('')
    scopus_id_entries = data['scopus_id'].values
    doi_entries = np.array(data['doi'].values)
    doi_entries = [f"{str(doi).replace('https://doi.org/', '')}" for doi in doi_entries]
    abstracts_entries = np.array(data['abstract'].values)
    for i in range(0,1535):
        if (str(abstracts_entries[i]) == ''):
            #url = f'https://api.elsevier.com/content/abstract/scopus_id/{str(scopus_id_entries[i]).replace(".0","")}'
            url = f'https://api.elsevier.com/content/search/scopus?query=DOI({doi_entries[i]})&view=COMPLETE&field=dc:title,dc:description,prism:doi&APIKey=b5131d0ad91b62931d30b30f2e2252df'

            # Making a get request 
            response = requests.get(url, headers={"Accept": "application/xml"})#, "X-ELS-APIKey": "fafe1117886418fdc64ee21c0347792a", "view": "COMPLETE", "field": "dc:title,dc:description,prism:doi"})  
            # print json content
            #print(url)
            if(response.status_code == 200):                   
                soup = BS(response.content, features="xml")
                
                if soup.find('dc:description') is not None:
                    print(soup.find('dc:description').text)
                    abstracts_entries[i] = ''
                    abstracts_entries[i] = soup.find('dc:description').text
                else:
                    print("not found")
                    print(response.content)

    data['abstract'] = abstracts_entries
    data.to_csv(f'{file}', index=False)
    
def get_abstract_from_scopusid(scopus_id):
    
    url = f"https://api.elsevier.com/content/abstract/scopus_id/{scopus_id}"
    response = requests.get(url, headers={"Accept": "application/xml", "X-ELS-APIKey": "b5131d0ad91b62931d30b30f2e2252df"})
    print(response.status_code)
    print(response.content)
    soup = BS(response.content, features="xml")
    abstract = "null"
    if soup.find('abstracts') is not None:
        abstract = soup.find('abstracts')
    print(abstract)


def check_screening(file1, file2):
    data1 = pd.read_csv(file1, encoding_errors='ignore')
    data2 = pd.read_csv(file2, encoding_errors='ignore')
    
    screenedDOI_entries = data2['doi'].values
    KY_entries = data2['KY Screening Decision'].values
    PA_entries = data2['PA Screening Decision'].values
    
    #create a screened dictionary doi as key, KY and PA screened results as values
    screened = {}
    #for i in range(0,30):
    for i in range(len(data2)):
        # print(screenedDOI_entries[i])
        screened[screenedDOI_entries[i]] = {"KY": KY_entries[i], "PA": PA_entries[i]}
        # print(screened[screenedDOI_entries[i]])
    
    newKY_entries = ['' for i in range(len(data1))]
    newPA_entries = ['' for i in range(len(data1))]
    DOI_entries = data1['doi'].values
    #for i in range(0,30):
    for i in range(len(data1)):
        currDOI = DOI_entries[i]
        if currDOI in screened:
            newKY_entries[i] = screened[currDOI]["KY"]
            newPA_entries[i] = screened[currDOI]["PA"]
        else:
            newKY_entries[i] = ""
            newPA_entries[i] = ""
        
        # print(newKY_entries[i])
        # print(newPA_entries[i])
        # print('***')
            
    data1['KY screen'] = newKY_entries
    data1['PA screen'] = newPA_entries
    data1.to_csv(f'{file1}', index = False)

def check_abstract(file1, file2):
    data1 = pd.read_csv(file1, encoding_errors='ignore')
    data2 = pd.read_csv(file2, encoding_errors='ignore')
    
    pastedDOI_entries = data2['doi'].values
    pastedAbstract_entries = data2['abstract'].values
    
    #create a screened dictionary doi as key, abstract as values
    pasted = {}
    #for i in range(0,30):
    for i in range(len(data2)):
        # print(screenedDOI_entries[i])
        pasted[pastedDOI_entries[i]] = pastedAbstract_entries[i]
        #print(pasted[pastedDOI_entries[i]])
    
    abstract_entries = data1['abstract'].values
    DOI_entries = data1['doi'].values
    #for i in range(0,30):
    for i in range(len(data1)):
        if(abstract_entries[i]==''):
            currDOI = DOI_entries[i]
            if currDOI in pasted:
                abstract_entries[i] = pasted[currDOI]
            else:
                abstract_entries[i] = ""
        
        # print(newKY_entries[i])
        # print(newPA_entries[i])
        # print('***')
            
    data1['abstract'] = abstract_entries
    data1.to_csv(f'{file1}', index = False)
#bib_remove_no_author('combined.bib', 'filters')
#get_csv_from_bib('combined_remove_no_author.bib', '0618new')
#get_doi_from_title('screening-articles-0618new.csv')
#scopus_id('screening-articles-0618new.csv')
#abstract_from_scopus('screening-articles-0618new.csv')

 

#cite_by_from_scopus('screening-articles-0618new_PA.csv')
#calculate_ACR('screening-articles-0618new_PA.csv', 2024.5)
#from_csv_bib_entries('screening-articles-0618new_PA.csv', 'screening-articles-newPA.bib')
    
#read_bib_entries('screening-articles-newPA.bib', '0619')
#read_bib_entries('combined.bib', 'filters')
#read_bib_entries('removeDup_output.bib', 'filters')
#read_bib_entries('deleted.bib', 'filters')
#read_bib_entries('springer_article.bib', 'filters')

#check_screening('screening-articles-0619.csv', 'Screening421-KYPA.csv')

#open_all_doi('Screening457_ACR1.5.csv')
open_all_doi('Screening457_ACR1.5.csv')
#get_bib_list_from_csv('springerlink_conferencepaper.csv')
#get_bib_list_from_csv('springerlink_article.csv')
#get_bib_list_from_csv('WebOfScience.csv')

""" filter_selected_papers()
open_urls('screening-articles-FINAL_discussed.xlsx')
compare_decision()
from_csv_bib_entries()
open_all_doi('screening-articles-final.csv')
read_bib_entries('screening-articles-new.bib')
get_bib_list_from_csv('screening-articles.csv')
get_doi_metadata("10.1109/ICEIT57125.2023.10107848")
combine_files('ScienceDirect_citations', 4, "bib") """