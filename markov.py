
import xlrd
import re
from collections import Counter
import random


file_location = "C:/Users/Taylor/Documents/c.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)

file_location_images = "C:/Users/Taylor/Documents/images.xlsx"
workbook_images = xlrd.open_workbook(file_location_images)
sheet_images = workbook_images.sheet_by_index(0)



#create list of all words

list = []

#iterate through all rows and columns in excel file, remove
#commas, slit sentences into lists of words, add each list
#to the main list

for col in range(sheet.ncols):
    for row in range(sheet.nrows):

        a = str(sheet.cell_value(row, col))
        if a and "http" not in a:
          a = a.replace("'", "")
          a = a.replace("’", "")
          a = a.replace("‘", "")
          a = re.sub("[^\w]", " ", str(a).lower()).split()
          list.append(a)

 # make list of images

images = []
for col in range(sheet_images.ncols):
    for row in range(sheet_images.nrows):
        a = str(sheet_images.cell_value(row, col))
        if a:
          images.append(a)

#make list of each individual word used

keys_list=[]

for sentence in list:
    for word in sentence:

        if word not in keys_list:
            keys_list.append(word)


finalwordlist = {}

firstwords = {}

#iterate through each word, and then iterate through the list of all sentences

for i in range(len(keys_list)):
    testword = keys_list[i]
    # add each word that follows the given word to proceedingwords[]
    # if it is the last word in the sentence, add None
    proceedingwords = []
    for sentence in list:
        for word in sentence:

            if word == testword:


                if sentence.index(word)+1 == len(sentence):
                    proceedingwords.append(None)
                else:
                    proceedingwords.append(list[list.index(sentence)][sentence.index(word)+1])

    #define "a" as a dictionary containing the total number occurances of that word
    #when it comes after testword
    size = len(proceedingwords)
    a = Counter(proceedingwords)
    #normalize to determine probablity
    for j in a:
        a[j] = float(a[j]/size)

    #append to dictionary the testword and the dictionary of its following words and their frequency
    finalwordlist.update({keys_list[i]: a})

#make list of words that have more than one possible following word, just to make sure
#sentences are more than one word long

for x,y in finalwordlist.items():

    if len([v for v in y]) != 1:
            firstwords.update({x: y})

#make number of words per sentence a random number, and number of sentences a random number

numwords = random.randrange(10, 20)
numsentences = random.randrange(3, 5)
#create list
sentence = []

for j in range(numsentences):
    #make random firstword the first word of the sentences
    sentence.append(random.choice([k for k in firstwords]))
    for i in range(numwords):
        #make list of possible next words to choose from
        if i == 0:
            #dont append "none" if it's the first word in a sentence
            words = [k for k in finalwordlist[sentence[len(sentence)-1]] if k is not None]
        else:
            words = [k for k in finalwordlist[sentence[len(sentence) - 1]]]
        #choose random word - honestly i couldn't figure out to choose it with the probability that
        #i calculated earlier but i'm too lazy to remove that from the code
        addedWord = random.choice(words)
        #if it adds None, concatonate a period to the last word and end the loop
        if addedWord is None:
            s1 = sentence[len(sentence) - 1] + "."
            sentence.remove(sentence[len(sentence) - 1])
            sentence.append(s1)
            break

        else:
            sentence.append(addedWord)
        #if it's the last word in the loop just add a period to the end
        if i == (numwords - 1):
            s1 = sentence[len(sentence) - 1] + "."
            sentence.remove(sentence[len(sentence) - 1])
            sentence.append(s1)
#join the list as string
print(' '.join(sentence))
#choose random image
print(random.choice(images))
