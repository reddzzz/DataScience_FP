import pandas as pd
import re
import string
import time
import os

# data = pd.read_excel("mydataset_summary.xlsx", nrows = 1000)
data=pd.read_excel(
     os.path.join("dataset.xlsx"), nrows = 1000,
     engine='openpyxl',
)

print(data.columns)

data = data.rename(columns={'content':'articles'})

data = data.filter(['human_summary','articles','theme'])
print(data.isnull().sum())
data.head()

contraction_expander = {
"ain't": "am not", "aren't": "are not", "can't": "cannot", "can't've": "cannot have", "'cause": "because", "could've": "could have", "couldn't": "could not",
"couldn't've": "could not have", "didn't": "did not", "doesn't": "does not", "don't": "do not", "hadn't": "had not", "hadn't've": "had not have", "hasn't": "has not",
"haven't": "have not", "he'd": "he would", "he'd've": "he would have","he'll": "he will", "he's": "he is", "how'd": "how did", "how'll": "how will","how's": "how is",
"i'd": "i would","i'll": "i will","i'm": "i am","i've": "i have", "isn't": "is not","it'd": "it would","it'll": "it will","it's": "it is","let's": "let us","ma'am": "madam",
"mayn't": "may not","might've": "might have","mightn't": "might not","must've": "must have","mustn't": "must not","needn't": "need not","oughtn't": "ought not","shan't": "shall not",
"sha'n't": "shall not","she'd": "she would", "she'll": "she will","she's": "she is","should've": "should have","shouldn't": "should not","that'd": "that would", "that's": "that is",
"there'd": "there had","there's": "there is","they'd": "they would","they'll": "they will", "they're": "they are","they've": "they have","wasn't": "was not","we'd": "we would",
"we'll": "we will","we're": "we are","we've": "we have","weren't": "were not","what'll": "what will","what're": "what are","what's": "what is", "what've": "what have",
"where'd": "where did","where's": "where is","who'll": "who will","who's": "who is", "won't": "will not","wouldn't": "would not","you'd": "you would","you'll": "you will",
"you're": "you are"
}



def cleaning(content):
    content = content.lower()
    # expand contractions
    expanded_form = []
    for i in content.split():
        if i in contraction_expander:
            expanded_form = expanded_form + [contraction_expander[i]]
        else:
            expanded_form = expanded_form + [i]
    content = ' '.join(expanded_form)

    content = re.sub(r"\([^()]*\)", "", content)  # remove words in brackets
    content = content.replace("_____", "")
    content = content.replace("■", "")
    content = content.replace("•","")

    punctuation = '''!()-[]|{};:'"\,<>/?@#$%^&*_~—“””'''  # punctuation
    for i in content:
        if i in punctuation:
            content = content.replace(i, "")

    content = content.replace("’s", "")
    content = content.replace("_______ •", "")

    content = re.sub('\s+', ' ', content).strip()
    content = re.sub("https*\S+", " ", content)

    return content


clean_summary = []
for summary in data.human_summary:
    clean_summary.append(cleaning(summary))

clean_text = []
for text in data.content:
    clean_text.append(cleaning(text))

data['clean_articles'] = clean_text
data['clean_summaries'] = clean_summary

from summarizer import Summarizer
data['computer_summary'] = " "

j = 0
start = time.time()
for i in range(len(data['clean_articles'])):
    model = Summarizer()
    body = data['clean_articles'][i]
    result = model(body, min_length=200)
    full = ''.join(result)
    data['computer_summary'][i] = full
    j = j + 1
    print("Completed: ", j)
end = time.time()
time_taken = (end-start)/60
print("Time Taken: ", time_taken, " minutes")

from rouge import Rouge

print("Calculating Rouge scores: ")

c,r = map(list, zip(*[[data['computer_summary'][i], data['human_summary'][i]] for i in range (len(data))]))

rouge = Rouge()

scores = rouge.get_scores(c,r,avg=True)
print(scores)

for i in range(len(data['clean_articles'])):
  print("Computer Generated Summary: ")
  print(i)
  print(data['computer_summary'][i])
  print("\n")


