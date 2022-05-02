###############################################################
##                                                           ##
##                       MORI AutoCoder                      ##
##                      ---------------                      ##
##                                                           ##
##  Python script for coding MORI discussions                ##
##                                                           ##
##  Current version [1.0]                                    ##
##                                                           ##
##  Please see document for full explanation on              ##
##  customizing the script for your own study                ##
##                                                           ##
###############################################################


###############################################################
#                           IMPORTS                           #
###############################################################

import xlwings #Links to Excel
import re #Need regex 
import tkinter
import sys
from tkinter import *

###############################################################
#                       MAIN CODEBLOCK                        #
###############################################################
def moriD1():
    path = file_path.get()
    string = transcript_entry.get("1.0", "end")
    pair_number = pair_entry.get()
    total_text1 = string
    wb = xlwings.Book(path) #Finds Excel Book
    Sheet1 = wb.sheets[4] #grabs correct sheet (starts at zero - pythonic)
    
    Sheet1.range('A2','A3').value = pair_number # in the gui, when you put in pair number this puts it in the excel sheet
    Sheet1.range('A4','A5').value = pair_number # have to do it like this, since range only works for two numbers, apparently
    Sheet1.range('A6','A7').value = pair_number
    Sheet1.range('A8','A9').value = pair_number
    Sheet1.range('A10','A11').value = pair_number
    
    Sheet1.range('B2','B3').value = 'C' #when the correct discussion is clicked, it auto populates the C/D/E/F
    Sheet1.range('B4','B5').value = 'C'
    Sheet1.range('B6','B7').value = 'C'
    Sheet1.range('B8','B9').value = 'C'
    Sheet1.range('B10','B11').value = 'C'
    
    total_text2 = total_text1.replace('!','.')
    total_text = total_text2.replace('?','.') #replace every question mark with a period to bypass regex, now unnecessary, but held due to simplicity

    valid = list('ABCDEFKMGT') + ['SAR','CAW','SDL','EWM','ERS']

    for w in re.findall('\w+', total_text):
        if w not in valid:
            total_text = re.sub(w, w.lower(), total_text)
    questions = total_text.split('~') #split each questions up into a list
    
    #Final Misinfo for each participant - across all questions for each discussion type
    Final_MisinfoP1 = []
    Final_MisinfoP2 = []
    
    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 1:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[0] #Which of the questions (list number) is being read
    if 'question 1' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by new paragraph
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.
        
        
        Correct_AnswerP1 = 'D'
        Correct_AnswerP2 = 'D'

        Wrong_AnswerP1 = ['A','B','C','E']

        Wrong_AnswerP2 = ['A','B','C','E']

        CWAB = ['A','B','C','E']
        
        dic = {'A':'red','B':'white','C':'yellow','D':'blue','E':'green'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[1]: #This determines who spoke first AND IS DIFFERENT FOR NON-QUESTION 1'S BECAUSE THE TRANSCRIPT DOES NOT START WITH '~' WHICH IS DELETED DURING THE SPLIT. THEREFORE THE FIRST ITEM IN THE LIST IS EMPTY FOR ALL NON-QUESTION 1 ITEMS
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(2, 17).value = '1'
            Sheet1.range(2, 18).value = '2'
        elif 'person 2' in Chat[1]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(2, 17).value = '2'
            Sheet1.range(2, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'red' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'white' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'yellow' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'blue' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'green' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
                


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'red' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                
            if 'person 1' in each_sentence and 'white' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                
            if 'person 1' in each_sentence and 'yellow' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                
            if 'person 1' in each_sentence and 'blue' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')

            if 'person 1' in each_sentence and 'green' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(2, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(2, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(2, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(2, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(2, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(2, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(2, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(2, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(2, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(2, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(2, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(2, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(2, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(2, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(2, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(2, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(2, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(2, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(2, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(2, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(2, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(2, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(2, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(2, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(2, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(2, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(2, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(2, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 11).value = '2' #Can be changed to be a different number here for different information

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(2, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(2, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 12).value = '2' #Can be changed to something else, to differentiate. 

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(2, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(2, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(2, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(2, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(2, 9).value = '3'
            Sheet1.range(2, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(2, 9).value = '4'
            Sheet1.range(2, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(2, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(2, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(2, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(2, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(2, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(2, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(2, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(2, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines
    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(2, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(2, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(2, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(2, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(2, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(2, 16).value = '3'
            
        
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 1 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 2 IS A FILLER
    #------------------------

    #------------------------
    # QUESTION 3 IS A FILLER
    #------------------------


    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 4:
    #-------------------------------------------------------------------------------------------------------------------
    
    txt = questions[3] #Which of the questions (list number) is bein read
    if 'question 4' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['A','B','D','E']

        Wrong_AnswerP2 = ['A','E','C','D']

        CWAB = ['A','D','E']
        
        dic = {'A':'sagrada familia','B':'leaning tower of pisa','C':'eiffel tower','D':'the london eye','E':'the taj mahal'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(8, 17).value = '1'
            Sheet1.range(8, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(8, 17).value = '2'
            Sheet1.range(8, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'sagrada' in each_sentence or 'person 2' in each_sentence and 'familia' in each_sentence or 'person 2' in each_sentence and 'sagrada familia' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                
            if 'person 2' in each_sentence and 'leaning' in each_sentence or 'person 2' in each_sentence and 'pisa' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'eiffel' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'london' in each_sentence or 'person 2' in each_sentence and 'london eye' in each_sentence or 'person 2' in each_sentence and 'eye' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'taj' in each_sentence or 'person 2' in each_sentence and 'mahal' in each_sentence or 'person 2' in each_sentence and 'taj mahal' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
    

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'sagrada' in each_sentence or 'person 1' in each_sentence and 'familia' in each_sentence or 'person 1' in each_sentence and 'sagrada familia' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                
            if 'person 1' in each_sentence and 'leaning' in each_sentence or 'person 1' in each_sentence and 'pisa' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'eiffel' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'london' in each_sentence or 'person 1' in each_sentence and 'london eye' in each_sentence or 'person 1' in each_sentence and 'eye' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'taj' in each_sentence or 'person 1' in each_sentence and 'mahal' in each_sentence or 'person 1' in each_sentence and 'taj mahal' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []
        check = any(item in CWAB for item in MisinfoP1)

        #Analysis
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(8, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(8, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(8, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(8, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(8, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(8, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(8, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(8, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(8, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(8, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(8, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(8, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(8, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(8, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(8, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(8, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(8, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(8, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(8, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(8, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(8, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(8, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(8, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(8, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(8, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(8, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(8, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(8, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(8, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(8, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(8, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(8, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(8, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(8, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(8, 9).value = '3'
            Sheet1.range(8, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(8, 9).value = '4'
            Sheet1.range(8, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(8, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(8, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(8, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(8, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(8, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(8, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(8, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(8, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(8, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(8, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(8, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(8, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(8, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(8, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    
    wb.save()
    
    

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 4 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 5:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[4] #Which of the questions (list number) is bein read
    if 'question 5' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','B','C','D']

        Wrong_AnswerP2 = ['A','B','C','D']

        CWAB = ['A','B','C','D']
        
        dic = {'A':'pens','B':'sunglasses','C':'watch','D':'necklace','E':'glove'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(4, 17).value = '1'
            Sheet1.range(4, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(4, 17).value = '2'
            Sheet1.range(4, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'pen' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'sunglasses' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'watch' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'necklace' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'glove' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'pen' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'sunglasses' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'watch' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'necklace' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'glove' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(4, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(4, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(4, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(4, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(4, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(4, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(4, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(4, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(4, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(4, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(4, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(4, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(4, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(4, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(4, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(4, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(4, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(4, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(4, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(4, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(4, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(4, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(4, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(4, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(4, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(4, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(4, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(4, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(4, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(4, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(4, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(4, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(4, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(4, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(4, 9).value = '3'
            Sheet1.range(4, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(4, 9).value = '4'
            Sheet1.range(4, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(4, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(4, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(4, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(4, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(4, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(4, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(4, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(4, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(4, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(4, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(4, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(4, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(4, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(4, 16).value = '3'
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)  
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 5 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 6:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[5] #Which of the questions (list number) is bein read
    if 'question 6' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['D','B','C','E']

        Wrong_AnswerP2 = ['A','D','C','E']

        CWAB = ['C','D','E'] #combined wrong asnwers for both - used in exposure
        
        dic = {'A':'metal spoon','B':'plastic fork','C':'metal knife','D':'plastic knife','E':'plastic spoon'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(10, 17).value = '1'
            Sheet1.range(10, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(10, 17).value = '2'
            Sheet1.range(10, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'metal spoon' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')

            if 'person 2' in each_sentence and 'plastic fork' in each_sentence or 'person 2' in each_sentence and 'fork' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'metal knife' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'plastic knife' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'plastic spoon' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'metal spoon' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'plastic fork' in each_sentence or 'person 1' in each_sentence and 'fork' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'metal knife' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'plastic knife' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'plastic spoon' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                
        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(10, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(10, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(10, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(10, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(10, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(10, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(10, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(10, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(10, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(10, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(10, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(10, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(10, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(10, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(10, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(10, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(10, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(10, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(10, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(10, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(10, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(10, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(10, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(10, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(10, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(10, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(10, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(10, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(10, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(10, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(10, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(10, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(10, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(10, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(10, 9).value = '3'
            Sheet1.range(10, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(10, 9).value = '4'
            Sheet1.range(10, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(10, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(10, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(10, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(10, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(10, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(10, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(10, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(10, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(10, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(10, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(10, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(10, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(10, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(10, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
        
    wb.save()


    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 6 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 7 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 8:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[7] #Which of the questions (list number) is bein read
    if 'question 8' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','B','C','D']

        Wrong_AnswerP2 = ['A','B','C','D']

        CWAB = ['A','B','C','D']
        
        dic = {'A':'14:00','B':'21:00','C':'8:30','D':'16:30','E':'9:30'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(3, 17).value = '1'
            Sheet1.range(3, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(3, 17).value = '2'
            Sheet1.range(3, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and '14:00' in each_sentence or 'person 2' in each_sentence and 'fourteen' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and '21:00' in each_sentence or 'person 2' in each_sentence and 'twenty-one' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and '8:30' in each_sentence or 'person 2' in each_sentence and 'eight-thirty' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and '16:30' in each_sentence or 'person 2' in each_sentence and 'sixteen-thirty' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and '9:30' in each_sentence or 'person 2' in each_sentence and 'nine-thirty' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
 

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and '14:00' in each_sentence or 'person 1' in each_sentence and 'fourteen' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and '21:00' in each_sentence or 'person 1' in each_sentence and 'twenty-one' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and '8:30' in each_sentence or 'person 1' in each_sentence and 'eight-thirty' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and '16:30' in each_sentence or 'person 1' in each_sentence and 'sixteen-thirty' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and '9:30' in each_sentence or 'person 1' in each_sentence and 'nine-thirty' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(3, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(3, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(3, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(3, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(3, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(3, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(3, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(3, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(3, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(3, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(3, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(3, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(3, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(3, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(3, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(3, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(3, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(3, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(3, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(3, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(3, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(3, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(3, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(3, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(3, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(3, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(3, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(3, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(3, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(3, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(3, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(3, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(3, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(3, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(3, 9).value = '3'
            Sheet1.range(3, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(3, 9).value = '4'
            Sheet1.range(3, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(3, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(3, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(3, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(3, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(3, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(3, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(3, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(3, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(3, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(3, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(3, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(3, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(3, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(3, 16).value = '3'
    
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
        
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 8 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 9:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[8] #Which of the questions (list number) is bein read
    if 'question 9' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['A','D','C','E']

        Wrong_AnswerP2 = ['A','B','D','E']

        CWAB = ['A','D','E']
        
        dic = {'A':'apple','B':'house','C':'tree','D':'helicopter','E':'truck'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(7, 17).value = '1'
            Sheet1.range(7, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(7, 17).value = '2'
            Sheet1.range(7, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'apple' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'house' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'tree' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'helicopter' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'truck' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'apple' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'house' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'tree' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'helicopter' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'truck' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(7, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(7, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(7, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(7, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(7, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(7, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(7, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(7, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(7, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(7, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(7, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(7, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(7, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(7, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(7, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(7, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(7, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(7, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(7, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(7, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(7, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(7, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(7, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(7, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(7, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(7, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(7, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(7, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(7, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(7, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(7, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(7, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(7, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(7, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(7, 9).value = '3'
            Sheet1.range(7, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(7, 9).value = '4'
            Sheet1.range(7, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(7, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(7, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(7, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(7, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(7, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(7, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(7, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(7, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(7, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(7, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(7, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(7, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(7, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(7, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 9 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 10 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 11:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[10] #Which of the questions (list number) is bein read
    if 'question 11' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'A'

        Wrong_AnswerP1 = ['D','B','C','E']

        Wrong_AnswerP2 = ['D','B','C','E']

        CWAB = ['B','C','D','E']
        
        dic = {'A':'cafe','B':'toilets','C':'laboratory','D':'library','E':'car park'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(5, 17).value = '1'
            Sheet1.range(5, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(5, 17).value = '2'
            Sheet1.range(5, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'cafe' in each_sentence or 'person 2' in each_sentence and 'café' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'toilets' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'laboratory' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'library' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'car' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'cafe' in each_sentence or 'person 1' in each_sentence and 'café' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'toilets' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'laboratory' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'library' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'car' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(5, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(5, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(5, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(5, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(5, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(5, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(5, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(5, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(5, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(5, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(5, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(5, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(5, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(5, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(5, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(5, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(5, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(5, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(5, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(5, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(5, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(5, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(5, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(5, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(5, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(5, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(5, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(5, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(5, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(5, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(5, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(5, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(5, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(5, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(5, 9).value = '3'
            Sheet1.range(5, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(5, 9).value = '4'
            Sheet1.range(5, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(5, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(5, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(5, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(5, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(5, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(5, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(5, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(5, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(5, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(5, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(5, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(5, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(5, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(5, 16).value = '3'
        
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 11 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 12:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[11] #Which of the questions (list number) is bein read
    if 'question 12' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','D','C','E']

        Wrong_AnswerP2 = ['A','B','C','D']

        CWAB = ['A','D','C']
        
        dic = {'A':'elephant','B':'tiger','C':'dog','D':'ostrich','E':'panther'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(9, 17).value = '1'
            Sheet1.range(9, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(9, 17).value = '2'
            Sheet1.range(9, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'elephant' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'tiger' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'dog' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'ostrich' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'panther' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'elephant' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'tiger' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'dog' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'ostrich' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'panther' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(9, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(9, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(9, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(9, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(9, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(9, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(9, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(9, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(9, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(9, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(9, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(9, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(9, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(9, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(9, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(9, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(9, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(9, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(9, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(9, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(9, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(9, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(9, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(9, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(9, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(9, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(9, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(9, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(9, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(9, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(9, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(9, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(9, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(9, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(9, 9).value = '3'
            Sheet1.range(9, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(9, 9).value = '4'
            Sheet1.range(9, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(9, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(9, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(9, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(9, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(9, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(9, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(9, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(9, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(9, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(9, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(9, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(9, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(9, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(9, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 12 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 13 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 14:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[13] #Which of the questions (list number) is bein read
    if 'question 14' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['A','D','C','E']

        Wrong_AnswerP2 = ['A','D','C','E']

        CWAB = ['A','C','D','E']
        
        dic = {'A':'river','B':'mountains','C':'rainforest','D':'pyramids','E':'beach'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(6, 17).value = '1'
            Sheet1.range(6, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(6, 17).value = '2'
            Sheet1.range(6, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'river' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'mountain' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'rainforest' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'pyramid' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'beach' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
  

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'river' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'mountain' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'rainforest' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'pyramid' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'beach' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(6, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(6, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(6, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(6, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(6, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(6, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(6, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(6, 13).value = '4'


    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(6, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(6, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(6, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(6, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(6, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(6, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(6, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(6, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(6, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(6, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(6, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(6, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(6, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(6, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(6, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(6, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(6, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(6, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(6, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(6, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(6, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(6, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(6, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(6, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(6, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(6, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(6, 9).value = '3'
            Sheet1.range(6, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(6, 9).value = '4'
            Sheet1.range(6, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(6, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(6, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(6, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(6, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(6, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(6, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(6, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(6, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(6, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(6, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(6, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(6, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(6, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(6, 16).value = '3'
    
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 14 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 15:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[14] #Which of the questions (list number) is bein read
    if 'question 15' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','B','D','E']

        Wrong_AnswerP2 = ['A','B','C','D']

        CWAB = ['A','B','D']
        
        dic = {'A':'wallet','B':'handbag','C':'cash','D':'bracelet','E':'credit card'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(11, 17).value = '1'
            Sheet1.range(11, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(11, 17).value = '2'
            Sheet1.range(11, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'wallet' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'handbag' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'cash' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'bracelet' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'card' in each_sentence or 'person 2' in each_sentence and 'credit' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'wallet' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'handbag' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'cash' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'bracelet' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'card' in each_sentence or 'person 1' in each_sentence and 'credit' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(11, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(11, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(11, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(11, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(11, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(11, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(11, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(11, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(11, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(11, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(11, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(11, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(11, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(11, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(11, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(11, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(11, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(11, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(11, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(11, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(11, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(11, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(11, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(11, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(11, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(11, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(11, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(11, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(11, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(11, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(11, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(11, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(11, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(11, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(11, 9).value = '3'
            Sheet1.range(11, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(11, 9).value = '4'
            Sheet1.range(11, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(11, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(11, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(11, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(11, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(11, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(11, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(11, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(11, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(11, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(11, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(11, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(11, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(11, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(11, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    
    #Final Exposure Results
    print(Final_MisinfoP1)
    print(Final_MisinfoP2)
    Sheet1.range(2, 19).value = len(Final_MisinfoP1)
    Sheet1.range(2, 20).value = len(Final_MisinfoP2)
    
    print('\n' + '~'*60 + '\n' + ' '*21 + 'ANALYSIS FINISHED PLEASE CHECK ERRORS AND EXCEL SHEET' + '\n' + '~'*60)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 15 FINISHED
    # --------------------------------------------------------------------------------------------------------
    
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
                # MORI
                # Discussion 2
                # Below
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------


def moriD2():
    path = file_path.get()
    string = transcript_entry.get("1.0", "end")
    pair_number = pair_entry.get()
    total_text1 = string
    wb = xlwings.Book(path) #Finds Excel Book
    Sheet1 = wb.sheets[4] #grabs correct sheet (starts at zero - pythonic)
    
    Sheet1.range('A2','A3').value = pair_number #pair number
    Sheet1.range('A4','A5').value = pair_number
    Sheet1.range('A6','A7').value = pair_number
    Sheet1.range('A8','A9').value = pair_number
    Sheet1.range('A10','A11').value = pair_number
    
    Sheet1.range('B2','B3').value = 'D' #when the correct discussion is clicked, it auto populates the C/D/E/F
    Sheet1.range('B4','B5').value = 'D'
    Sheet1.range('B6','B7').value = 'D'
    Sheet1.range('B8','B9').value = 'D'
    Sheet1.range('B10','B11').value = 'D'
    
    total_text2 = total_text1.replace('!','.')
    total_text = total_text2.replace('?','.') #replace every question mark with a period to bypass regex

    valid = list('ABCDEFKMGT') + ['SAR','CAW','SDL','EWM','ERS']

    for w in re.findall('\w+', total_text):
        if w not in valid:
            total_text = re.sub(w, w.lower(), total_text)
    questions = total_text.split('~') #split each questions up into a list
    
    #Final Misinfo for each participant - across all questions for each discussion type
    Final_MisinfoP1 = []
    Final_MisinfoP2 = []

    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 1:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[0] #Which of the questions (list number) is bein read
    if 'question 1' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'D'
        Correct_AnswerP2 = 'D'

        Wrong_AnswerP1 = ['A','B','C','E']

        Wrong_AnswerP2 = ['A','B','C','E']

        CWAB = ['A','B','C','E']
        
        dic = {'A':'frying pan','B':'baking tray','C':'casserole dish','D':'sieve','E':'mixing bowl'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[1]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(2, 17).value = '1'
            Sheet1.range(2, 18).value = '2'
        elif 'person 2' in Chat[1]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(2, 17).value = '2'
            Sheet1.range(2, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'frying pan' in each_sentence or 'person 2' in each_sentence and 'frying' in each_sentence or 'person 2' in each_sentence and 'pan' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'baking tray' in each_sentence or 'person 2' in each_sentence and 'tray' in each_sentence or 'person 2' in each_sentence and 'baking' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'casserole dish' in each_sentence or 'person 2' in each_sentence and 'casserole' in each_sentence or 'person 2' in each_sentence and 'dish' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'sieve' in each_sentence or 'person 2' in each_sentence and 'strainer' in each_sentence or 'person 2' in each_sentence and 'sieve/strainer' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'mixing bowl' in each_sentence or 'person 2' in each_sentence and 'mixing' in each_sentence or 'person 2' in each_sentence and 'bowl' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'frying pan' in each_sentence or 'person 1' in each_sentence and 'frying' in each_sentence or 'person 1' in each_sentence and 'pan' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'baking tray' in each_sentence or 'person 1' in each_sentence and 'tray' in each_sentence or 'person 1' in each_sentence and 'baking' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'casserole dish' in each_sentence or 'person 1' in each_sentence and 'casserole' in each_sentence or 'person 1' in each_sentence and 'dish' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'sieve' in each_sentence or 'person 1' in each_sentence and 'strainer' in each_sentence or 'person 1' in each_sentence and 'sieve/strainer' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'mixing bowl' in each_sentence or 'person 1' in each_sentence and 'mixing' in each_sentence or 'person 1' in each_sentence and 'bowl' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(2, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(2, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(2, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(2, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(2, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(2, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(2, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(2, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')


        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(2, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(2, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(2, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(2, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(2, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(2, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(2, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(2, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(2, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(2, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(2, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(2, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(2, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(2, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(2, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(2, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(2, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(2, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(2, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(2, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(2, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(2, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(2, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(2, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(2, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(2, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(2, 9).value = '3'
            Sheet1.range(2, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(2, 9).value = '4'
            Sheet1.range(2, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(2, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(2, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(2, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(2, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(2, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(2, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(2, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(2, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines
    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(2, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(2, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(2, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(2, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(2, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(2, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 1 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 2 IS A FILLER
    #------------------------

    #------------------------
    # QUESTION 3 IS A FILLER
    #------------------------


    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 4:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[3] #Which of the questions (list number) is bein read
    if 'question 4' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'D'

        Wrong_AnswerP1 = ['E','B','D','C']

        Wrong_AnswerP2 = ['A','E','C','B']

        CWAB = ['E','B','C']
        
        dic = {'A':'brown','B':'green','C':'grey','D':'purple','E':'black'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(8, 17).value = '1'
            Sheet1.range(8, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(8, 17).value = '2'
            Sheet1.range(8, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'brown' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'green' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'grey' in each_sentence or 'person 2' in each_sentence and 'gray' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'purple' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'black' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'brown' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'green' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'grey' in each_sentence or 'person 1' in each_sentence and 'gray' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'purple' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'black' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []
        check = any(item in CWAB for item in MisinfoP1)

        #Analysis
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(8, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(8, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(8, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(8, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(8, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(8, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(8, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(8, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(8, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(8, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(8, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(8, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(8, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(8, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(8, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(8, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(8, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(8, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(8, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(8, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(8, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(8, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(8, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(8, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(8, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(8, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(8, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(8, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(8, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(8, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(8, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(8, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(8, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(8, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(8, 9).value = '3'
            Sheet1.range(8, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(8, 9).value = '4'
            Sheet1.range(8, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(8, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(8, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(8, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(8, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(8, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(8, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(8, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(8, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(8, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(8, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(8, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(8, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(8, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(8, 16).value = '3'
    
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 4 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 5:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[4] #Which of the questions (list number) is bein read
    if 'question 5' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'A'

        Wrong_AnswerP1 = ['E','B','C','D']

        Wrong_AnswerP2 = ['E','B','C','D']

        CWAB = ['E','B','C','D']
        
        dic = {'A':'blue','B':'black','C':'pink','D':'white','E':'brown'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(4, 17).value = '1'
            Sheet1.range(4, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(4, 17).value = '2'
            Sheet1.range(4, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'blue' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'black' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'pink' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'white' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'brown' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'blue' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'black' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'pink' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'white' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'brown' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(4, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(4, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(4, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(4, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(4, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(4, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(4, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(4, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(4, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(4, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(4, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(4, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(4, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(4, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(4, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(4, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(4, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(4, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(4, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(4, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(4, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(4, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(4, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(4, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(4, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(4, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(4, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(4, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(4, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(4, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(4, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(4, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(4, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(4, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(4, 9).value = '3'
            Sheet1.range(4, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(4, 9).value = '4'
            Sheet1.range(4, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(4, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(4, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(4, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(4, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(4, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(4, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(4, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(4, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(4, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(4, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(4, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(4, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(4, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(4, 16).value = '3'
    
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 5 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 6:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[5] #Which of the questions (list number) is bein read
    if 'question 6' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'D'

        Wrong_AnswerP1 = ['D','B','C','A']

        Wrong_AnswerP2 = ['A','B','C','E']

        CWAB = ['C','A','B'] #combined wrong asnwers for both - used in exposure
        
        dic = {'A':'hard drive','B':'phone','C':'microphone','D':'tablet','E':'laptop'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(10, 17).value = '1'
            Sheet1.range(10, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(10, 17).value = '2'
            Sheet1.range(10, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'hard' in each_sentence or 'person 2' in each_sentence and 'drive' in each_sentence or 'person 2' in each_sentence and 'harddrive' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'phone' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'microphone' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'tablet' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'laptop' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'hard' in each_sentence or 'person 1' in each_sentence and 'drive' in each_sentence or 'person 1' in each_sentence and 'harddrive' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'phone' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'microphone' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'tablet' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'laptop' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(10, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(10, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(10, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(10, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(10, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(10, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(10, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(10, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(10, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(10, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(10, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(10, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(10, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(10, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(10, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(10, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(10, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(10, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(10, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(10, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(10, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(10, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(10, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(10, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(10, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(10, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(10, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(10, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(10, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(10, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(10, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(10, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(10, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(10, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(10, 9).value = '3'
            Sheet1.range(10, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(10, 9).value = '4'
            Sheet1.range(10, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(10, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(10, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(10, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(10, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(10, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(10, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(10, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(10, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(10, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(10, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(10, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(10, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(10, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(10, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()


    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 6 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 7 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 8:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[7] #Which of the questions (list number) is bein read
    if 'question 8' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','B','C','D']

        Wrong_AnswerP2 = ['A','B','C','D']

        CWAB = ['A','B','C','D']
        
        dic = {'A':'F','B':'K','C':'M','D':'G','E':'T'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(3, 17).value = '1'
            Sheet1.range(3, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(3, 17).value = '2'
            Sheet1.range(3, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'F' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'K' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'M' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'G' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'T' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'F' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'K' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'M' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'G' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'T' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(3, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(3, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(3, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(3, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(3, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(3, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(3, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(3, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(3, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(3, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(3, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(3, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(3, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(3, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(3, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(3, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(3, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(3, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(3, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(3, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(3, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(3, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(3, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(3, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(3, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(3, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(3, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(3, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(3, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(3, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(3, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(3, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(3, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(3, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(3, 9).value = '3'
            Sheet1.range(3, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(3, 9).value = '4'
            Sheet1.range(3, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(3, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(3, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(3, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(3, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(3, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(3, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(3, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(3, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(3, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(3, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(3, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(3, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(3, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(3, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 8 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 9:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[8] #Which of the questions (list number) is bein read
    if 'question 9' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','D','B','E']

        Wrong_AnswerP2 = ['A','B','D','C']

        CWAB = ['A','D','B']
        
        dic = {'A':'SAR','B':'CAW','C':'SDL','D':'EWM','E':'ERS'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(7, 17).value = '1'
            Sheet1.range(7, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(7, 17).value = '2'
            Sheet1.range(7, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'SAR' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'CAW' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'SDL' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'EWM' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'ERS' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'SAR' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'CAW' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'SDL' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'EWM' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'ERS' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(7, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(7, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(7, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(7, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(7, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(7, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(7, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(7, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(7, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(7, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(7, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(7, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(7, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(7, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(7, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(7, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(7, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(7, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(7, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(7, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(7, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(7, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(7, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(7, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(7, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(7, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(7, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(7, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(7, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(7, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(7, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(7, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(7, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(7, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(7, 9).value = '3'
            Sheet1.range(7, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(7, 9).value = '4'
            Sheet1.range(7, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(7, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(7, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(7, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(7, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(7, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(7, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(7, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(7, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(7, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(7, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(7, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(7, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(7, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(7, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 9 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 10 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 11:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[10] #Which of the questions (list number) is bein read
    if 'question 11' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['D','B','A','E']

        Wrong_AnswerP2 = ['D','B','A','E']

        CWAB = ['B','A','D','E']
        
        dic = {'A':'jug/pitcher','B':'wine glass','C':'mug','D':'soda can','E':'flask'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(5, 17).value = '1'
            Sheet1.range(5, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(5, 17).value = '2'
            Sheet1.range(5, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'jug' in each_sentence or 'person 2' in each_sentence and 'pitcher' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'wine' in each_sentence or 'person 2' in each_sentence and 'glass' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'mug' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'soda' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'flask' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'jug' in each_sentence or 'person 1' in each_sentence and 'pitcher' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'wine' in each_sentence or 'person 1' in each_sentence and 'glass' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'mug' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'soda' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'flask' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(5, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(5, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(5, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(5, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(5, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(5, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(5, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(5, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(5, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(5, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(5, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(5, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(5, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(5, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(5, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(5, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(5, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(5, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(5, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(5, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(5, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(5, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(5, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(5, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(5, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(5, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(5, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(5, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(5, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(5, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(5, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(5, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(5, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(5, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(5, 9).value = '3'
            Sheet1.range(5, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(5, 9).value = '4'
            Sheet1.range(5, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(5, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(5, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(5, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(5, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(5, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(5, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(5, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(5, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(5, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(5, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(5, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(5, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(5, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(5, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 11 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 12:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[11] #Which of the questions (list number) is bein read
    if 'question 12' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['A','D','C','E']

        Wrong_AnswerP2 = ['A','B','E','D']

        CWAB = ['A','D','E']
        
        dic = {'A':'no food or drink','B':'no smoking','C':'no mobile phones','D':'no cameras','E':'no animals'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(9, 17).value = '1'
            Sheet1.range(9, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(9, 17).value = '2'
            Sheet1.range(9, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'food' in each_sentence or 'person 2' in each_sentence and 'no food or drink' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'smoking' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'mobile' in each_sentence or 'person 2' in each_sentence and 'phone' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'camera' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'animal' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'food' in each_sentence or 'person 1' in each_sentence and 'no food or drink' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'smoking' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'mobile' in each_sentence or 'person 1' in each_sentence and 'phone' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'camera' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'animal' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(9, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(9, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(9, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(9, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(9, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(9, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(9, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(9, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(9, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(9, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(9, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(9, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(9, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(9, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(9, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(9, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(9, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(9, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(9, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(9, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(9, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(9, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(9, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(9, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(9, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(9, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(9, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(9, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(9, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(9, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(9, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(9, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(9, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(9, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(9, 9).value = '3'
            Sheet1.range(9, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(9, 9).value = '4'
            Sheet1.range(9, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(9, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(9, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(9, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(9, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(9, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(9, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(9, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(9, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(9, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(9, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(9, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(9, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(9, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(9, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 12 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 13 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 14:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[13] #Which of the questions (list number) is bein read
    if 'question 14' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['A','D','C','E']

        Wrong_AnswerP2 = ['A','D','C','E']

        CWAB = ['A','C','D','E']
        
        dic = {'A':'nothing, the screen is off','B':'abstract pattern','C':'car advert','D':'news channel','E':'timetable'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(6, 17).value = '1'
            Sheet1.range(6, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(6, 17).value = '2'
            Sheet1.range(6, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'nothing' in each_sentence or 'person 2' in each_sentence and 'off.' in each_sentence or 'person 2' in each_sentence and 'off ' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'abstract' in each_sentence or 'person 2' in each_sentence and 'pattern' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'car' in each_sentence or 'person 2' in each_sentence and 'advert' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'news' in each_sentence or 'person 2' in each_sentence and 'channel' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'time' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'nothing' in each_sentence or 'person 1' in each_sentence and 'off.' in each_sentence or 'person 2' in each_sentence and 'off ' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'abstract' in each_sentence or 'person 1' in each_sentence and 'pattern' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'car' in each_sentence or 'person 1' in each_sentence and 'advert' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'news' in each_sentence or 'person 1' in each_sentence and 'channel' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'time' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(6, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(6, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(6, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(6, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(6, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(6, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(6, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(6, 13).value = '4'


    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(6, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(6, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(6, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(6, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(6, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(6, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(6, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(6, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(6, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(6, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(6, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(6, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(6, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(6, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(6, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(6, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(6, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(6, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(6, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(6, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(6, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(6, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(6, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(6, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(6, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(6, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(6, 9).value = '3'
            Sheet1.range(6, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(6, 9).value = '4'
            Sheet1.range(6, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(6, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(6, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(6, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(6, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(6, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(6, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(6, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(6, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(6, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(6, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(6, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(6, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(6, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(6, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 14 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 15:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[14] #Which of the questions (list number) is bein read
    if 'question 15' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'D'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['A','B','C','E']

        Wrong_AnswerP2 = ['A','B','E','D']

        CWAB = ['A','B','E']
        
        dic = {'A':'blender','B':'kettle','C':'toaster','D':'microwave','E':'slow cooker'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(11, 17).value = '1'
            Sheet1.range(11, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(11, 17).value = '2'
            Sheet1.range(11, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'blender' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'kettle' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'toaster' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'microwave' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'slow' in each_sentence or 'person 2' in each_sentence and 'cooker' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'blender' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'kettle' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'toaster' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'microwave' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'slow' in each_sentence or 'person 1' in each_sentence and 'cooker' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(11, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(11, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(11, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(11, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(11, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(11, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(11, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(11, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(11, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(11, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(11, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(11, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(11, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(11, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(11, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(11, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(11, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(11, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(11, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(11, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(11, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(11, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(11, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(11, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(11, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(11, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(11, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(11, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(11, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(11, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(11, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(11, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(11, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(11, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(11, 9).value = '3'
            Sheet1.range(11, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(11, 9).value = '4'
            Sheet1.range(11, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(11, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(11, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(11, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(11, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(11, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(11, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(11, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(11, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(11, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(11, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(11, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(11, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(11, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(11, 16).value = '3'

    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    #Final Exposure Results
    print(Final_MisinfoP1)
    print(Final_MisinfoP2)
    Sheet1.range(2, 19).value = len(Final_MisinfoP1)
    Sheet1.range(2, 20).value = len(Final_MisinfoP2)
    
    print('\n' + '~'*60 + '\n' + ' '*21 + 'ANALYSIS FINISHED PLEASE CHECK ERRORS AND EXCEL SHEET' + '\n' + '~'*60) # Aesthetic Break between Questions (Odd Spacing - can be one line)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 15 FINISHED
    # --------------------------------------------------------------------------------------------------------
    
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
                # MORI
                # Discussion 3
                # Below
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
    
def moriD3():
    path = file_path.get()
    string = transcript_entry.get("1.0", "end")
    pair_number = pair_entry.get()
    total_text1 = string
    wb = xlwings.Book(path) #Finds Excel Book
    Sheet1 = wb.sheets[4] #grabs correct sheet (starts at zero - pythonic)
    
    Sheet1.range('A2','A3').value = pair_number #pair number
    Sheet1.range('A4','A5').value = pair_number
    Sheet1.range('A6','A7').value = pair_number
    Sheet1.range('A8','A9').value = pair_number
    Sheet1.range('A10','A11').value = pair_number
    
    Sheet1.range('B2','B3').value = 'E' #when the correct discussion is clicked, it auto populates the C/D/E/F
    Sheet1.range('B4','B5').value = 'E'
    Sheet1.range('B6','B7').value = 'E'
    Sheet1.range('B8','B9').value = 'E'
    Sheet1.range('B10','B11').value = 'E'
    
    total_text2 = total_text1.replace('!','.')
    total_text = total_text2.replace('?','.') #replace every question mark with a period to bypass regex

    valid = list('ABCDEFKMGT') + ['SAR','CAW','SDL','EWM','ERS']

    for w in re.findall('\w+', total_text):
        if w not in valid:
            total_text = re.sub(w, w.lower(), total_text)
    questions = total_text.split('~') #split each questions up into a list
    
    #Final Misinfo for each participant - across all questions for each discussion type
    Final_MisinfoP1 = []
    Final_MisinfoP2 = []

    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 1:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[0] #Which of the questions (list number) is bein read
    if 'question 1' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['A','B','D','E']

        Wrong_AnswerP2 = ['A','B','D','E']

        CWAB = ['A','B','D','E']
        
        dic = {'A':'apple','B':'house','C':'tree','D':'helicopter','E':'truck'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[1]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(2, 17).value = '1'
            Sheet1.range(2, 18).value = '2'
        elif 'person 2' in Chat[1]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(2, 17).value = '2'
            Sheet1.range(2, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'apple' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'house' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'tree' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'helicopter' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'truck' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'apple' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'house' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'tree' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'helicopter' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'truck' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(2, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(2, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(2, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(2, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(2, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(2, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(2, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(2, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')


        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(2, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(2, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(2, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(2, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(2, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(2, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(2, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(2, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(2, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(2, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(2, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(2, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(2, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(2, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(2, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(2, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(2, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(2, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(2, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(2, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(2, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(2, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(2, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(2, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(2, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(2, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(2, 9).value = '3'
            Sheet1.range(2, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(2, 9).value = '4'
            Sheet1.range(2, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(2, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(2, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(2, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(2, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(2, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(2, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(2, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(2, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines
    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(2, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(2, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(2, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(2, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(2, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(2, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 1 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 2 IS A FILLER
    #------------------------

    #------------------------
    # QUESTION 3 IS A FILLER
    #------------------------


    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 4:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[3] #Which of the questions (list number) is bein read
    if 'question 4' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','B','D','E']

        Wrong_AnswerP2 = ['A','D','C','B']

        CWAB = ['A','B','D']
        
        dic = {'A':'14:00','B':'21:00','C':'8:30','D':'16:30','E':'9:30'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(8, 17).value = '1'
            Sheet1.range(8, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(8, 17).value = '2'
            Sheet1.range(8, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and '14:00' in each_sentence or 'person 2' in each_sentence and 'fourteen' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and '21:00' in each_sentence or 'person 2' in each_sentence and 'twenty-one' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and '8:30' in each_sentence or 'person 2' in each_sentence and 'eight-thirty' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and '16:30' in each_sentence or 'person 2' in each_sentence and 'sixteen-thirty' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and '9:30' in each_sentence or 'person 2' in each_sentence and 'nine-thirty' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and '14:00' in each_sentence or 'person 1' in each_sentence and 'fourteen' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and '21:00' in each_sentence or 'person 1' in each_sentence and 'twenty-one' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and '8:30' in each_sentence or 'person 1' in each_sentence and 'eight-thirty' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and '16:30' in each_sentence or 'person 1' in each_sentence and 'sixteen-thirty' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and '9:30' in each_sentence or 'person 1' in each_sentence and 'nine-thirty' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []
        check = any(item in CWAB for item in MisinfoP1)

        #Analysis
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(8, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(8, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(8, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(8, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(8, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(8, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(8, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(8, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(8, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(8, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(8, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(8, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(8, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(8, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(8, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(8, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(8, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(8, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(8, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(8, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(8, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(8, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(8, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(8, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(8, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(8, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(8, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(8, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(8, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(8, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(8, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(8, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(8, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(8, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(8, 9).value = '3'
            Sheet1.range(8, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(8, 9).value = '4'
            Sheet1.range(8, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(8, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(8, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(8, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(8, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(8, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(8, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(8, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(8, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(8, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(8, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(8, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(8, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(8, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(8, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 4 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 5:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[4] #Which of the questions (list number) is bein read
    if 'question 5' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['E','D','C','A']

        Wrong_AnswerP2 = ['E','D','C','A']

        CWAB = ['E','D','C','A']
        
        dic = {'A':'elephant','B':'tiger','C':'dog','D':'ostrich','E':'panther'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(4, 17).value = '1'
            Sheet1.range(4, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(4, 17).value = '2'
            Sheet1.range(4, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'elephant' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'tiger' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'dog' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'ostrich' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'panther' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'elephant' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'tiger' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'dog' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'ostrich' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'panther' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(4, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(4, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(4, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(4, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(4, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(4, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(4, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(4, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(4, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(4, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(4, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(4, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(4, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(4, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(4, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(4, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(4, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(4, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(4, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(4, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(4, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(4, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(4, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(4, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(4, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(4, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(4, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(4, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(4, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(4, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(4, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(4, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(4, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(4, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(4, 9).value = '3'
            Sheet1.range(4, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(4, 9).value = '4'
            Sheet1.range(4, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(4, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(4, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(4, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(4, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(4, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(4, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(4, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(4, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(4, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(4, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(4, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(4, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(4, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(4, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 5 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 6:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[5] #Which of the questions (list number) is bein read
    if 'question 6' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'A'

        Wrong_AnswerP1 = ['D','E','C','A']

        Wrong_AnswerP2 = ['D','B','C','E']

        CWAB = ['C','D','E'] #combined wrong asnwers for both - used in exposure
        
        dic = {'A':'cafe','B':'toilets','C':'laboratory','D':'library','E':'car park'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(10, 17).value = '1'
            Sheet1.range(10, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(10, 17).value = '2'
            Sheet1.range(10, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'cafe' in each_sentence or 'person 2' in each_sentence and 'café' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'toilet' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'laboratory' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'library' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'car' in each_sentence or 'person 2' in each_sentence and 'park' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'cafe' in each_sentence or 'person 1' in each_sentence and 'café' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'toilet' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'laboratory' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'library' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'car' in each_sentence or 'person 1' in each_sentence and 'park' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(10, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(10, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(10, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(10, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(10, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(10, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(10, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(10, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(10, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(10, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(10, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(10, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(10, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(10, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(10, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(10, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(10, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(10, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(10, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(10, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(10, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(10, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(10, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(10, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(10, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(10, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(10, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(10, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(10, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(10, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(10, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(10, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(10, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(10, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(10, 9).value = '3'
            Sheet1.range(10, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(10, 9).value = '4'
            Sheet1.range(10, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(10, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(10, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(10, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(10, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(10, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(10, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(10, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(10, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(10, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(10, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(10, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(10, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(10, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(10, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()


    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 6 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 7 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 8:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[7] #Which of the questions (list number) is bein read
    if 'question 8' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['A','E','C','D']

        Wrong_AnswerP2 = ['A','E','C','D']

        CWAB = ['A','E','C','D']
        
        dic = {'A':'sagrada familia','B':'leaning tower of pisa','C':'eiffel tower','D':'the london eye','E':'the taj mahal'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(3, 17).value = '1'
            Sheet1.range(3, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(3, 17).value = '2'
            Sheet1.range(3, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'sagrada' in each_sentence or 'person 2' in each_sentence and 'familia' in each_sentence or 'person 2' in each_sentence and 'sagrada familia' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'leaning' in each_sentence or 'person 2' in each_sentence and 'pisa' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'eiffel' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'london' in each_sentence or 'person 2' in each_sentence and 'london eye' in each_sentence or 'person 2' in each_sentence and 'eye' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'taj' in each_sentence or 'person 2' in each_sentence and 'mahal' in each_sentence or 'person 2' in each_sentence and 'taj mahal' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'sagrada' in each_sentence or 'person 1' in each_sentence and 'familia' in each_sentence or 'person 1' in each_sentence and 'sagrada familia' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'leaning' in each_sentence or 'person 1' in each_sentence and 'pisa' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'eiffel' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'london' in each_sentence or 'person 1' in each_sentence and 'london eye' in each_sentence or 'person 1' in each_sentence and 'eye' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'taj' in each_sentence or 'person 1' in each_sentence and 'mahal' in each_sentence or 'person 1' in each_sentence and 'taj mahal' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(3, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(3, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(3, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(3, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(3, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(3, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(3, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(3, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(3, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(3, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(3, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(3, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(3, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(3, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(3, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(3, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(3, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(3, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(3, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(3, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(3, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(3, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(3, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(3, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(3, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(3, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(3, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(3, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(3, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(3, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(3, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(3, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(3, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(3, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(3, 9).value = '3'
            Sheet1.range(3, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(3, 9).value = '4'
            Sheet1.range(3, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(3, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(3, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(3, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(3, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(3, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(3, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(3, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(3, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(3, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(3, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(3, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(3, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(3, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(3, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 8 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 9:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[8] #Which of the questions (list number) is bein read
    if 'question 9' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'D'

        Wrong_AnswerP1 = ['C','D','B','E']

        Wrong_AnswerP2 = ['A','B','E','C']

        CWAB = ['C','E','B']
        
        dic = {'A':'red','B':'white','C':'yellow','D':'blue','E':'green'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(7, 17).value = '1'
            Sheet1.range(7, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(7, 17).value = '2'
            Sheet1.range(7, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'red' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'white' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'yellow' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'blue' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'green' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
  

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'red' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                
            if 'person 1' in each_sentence and 'white' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                
            if 'person 1' in each_sentence and 'yellow' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                
            if 'person 1' in each_sentence and 'blue' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                
            if 'person 1' in each_sentence and 'green' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(7, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(7, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(7, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(7, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(7, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(7, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(7, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(7, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(7, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(7, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(7, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(7, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(7, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(7, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(7, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(7, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(7, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(7, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(7, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(7, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(7, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(7, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(7, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(7, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(7, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(7, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(7, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(7, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(7, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(7, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(7, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(7, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(7, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(7, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(7, 9).value = '3'
            Sheet1.range(7, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(7, 9).value = '4'
            Sheet1.range(7, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(7, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(7, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(7, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(7, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(7, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(7, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(7, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(7, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(7, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(7, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(7, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(7, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(7, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(7, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 9 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 10 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 11:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[10] #Which of the questions (list number) is bein read
    if 'question 11' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['D','C','A','E']

        Wrong_AnswerP2 = ['D','C','A','E']

        CWAB = ['C','A','D','E']
        
        dic = {'A':'metal spoon','B':'plastic fork','C':'metal knife','D':'plastic knife','E':'plastic spoon'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(5, 17).value = '1'
            Sheet1.range(5, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(5, 17).value = '2'
            Sheet1.range(5, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'metal spoon' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'plastic fork' in each_sentence or 'person 2' in each_sentence and 'fork' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'metal knife' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'plastic knife' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'plastic spoon' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'metal spoon' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'plastic fork' in each_sentence or 'person 1' in each_sentence and 'fork' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'metal knife' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'plastic knife' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'plastic spoon' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(5, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(5, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(5, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(5, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(5, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(5, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(5, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(5, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(5, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(5, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(5, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(5, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(5, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(5, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(5, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(5, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(5, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(5, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(5, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(5, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(5, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(5, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(5, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(5, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(5, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(5, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(5, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(5, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(5, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(5, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(5, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(5, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(5, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(5, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(5, 9).value = '3'
            Sheet1.range(5, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(5, 9).value = '4'
            Sheet1.range(5, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(5, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(5, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(5, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(5, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(5, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(5, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(5, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(5, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(5, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(5, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(5, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(5, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(5, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(5, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 11 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 12:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[11] #Which of the questions (list number) is bein read
    if 'question 12' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['A','D','C','B']

        Wrong_AnswerP2 = ['A','B','E','D']

        CWAB = ['A','B','D']
        
        dic = {'A':'pens','B':'sunglasses','C':'watch','D':'necklace','E':'glove'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(9, 17).value = '1'
            Sheet1.range(9, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(9, 17).value = '2'
            Sheet1.range(9, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'pen' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'sunglasses' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'watch' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'necklace' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'glove' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
 

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'pen' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'sunglasses' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'watch' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'necklace' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'glove' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(9, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(9, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(9, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(9, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(9, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(9, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(9, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(9, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(9, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(9, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(9, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(9, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(9, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(9, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(9, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(9, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(9, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(9, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(9, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(9, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(9, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(9, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(9, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(9, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(9, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(9, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(9, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(9, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(9, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(9, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(9, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(9, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(9, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(9, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(9, 9).value = '3'
            Sheet1.range(9, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(9, 9).value = '4'
            Sheet1.range(9, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(9, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(9, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(9, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(9, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(9, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(9, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(9, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(9, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(9, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(9, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(9, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(9, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(9, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(9, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 12 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 13 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 14:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[13] #Which of the questions (list number) is bein read
    if 'question 14' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','D','C','B']

        Wrong_AnswerP2 = ['A','D','C','B']

        CWAB = ['A','C','D','B']
        
        dic = {'A':'wallet','B':'handbag','C':'cash','D':'bracelet','E':'credit card'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(6, 17).value = '1'
            Sheet1.range(6, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(6, 17).value = '2'
            Sheet1.range(6, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'wallet' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'handbag' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'cash' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'bracelet' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'card' in each_sentence or 'person 2' in each_sentence and 'credit' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'wallet' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'handbag' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'cash' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'bracelet' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'card' in each_sentence or 'person 1' in each_sentence and 'credit' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(6, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(6, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(6, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(6, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(6, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(6, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(6, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(6, 13).value = '4'


    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(6, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(6, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(6, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(6, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(6, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(6, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(6, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(6, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(6, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(6, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(6, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(6, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(6, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(6, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(6, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(6, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(6, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(6, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(6, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(6, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(6, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(6, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(6, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(6, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(6, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(6, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(6, 9).value = '3'
            Sheet1.range(6, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(6, 9).value = '4'
            Sheet1.range(6, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(6, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(6, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(6, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(6, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(6, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(6, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(6, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(6, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(6, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(6, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(6, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(6, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(6, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(6, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 14 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 15:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[14] #Which of the questions (list number) is bein read
    if 'question 15' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'B'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','D','C','E']

        Wrong_AnswerP2 = ['A','B','C','D']

        CWAB = ['A','C','D']
        
        dic = {'A':'river','B':'mountains','C':'rainforest','D':'pyramids','E':'beach'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(11, 17).value = '1'
            Sheet1.range(11, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(11, 17).value = '2'
            Sheet1.range(11, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'river' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'mountain' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'rainforest' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'pyramid' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'beach' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
                   

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'river' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'mountain' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'rainforest' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'pyramid' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'beach' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(11, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(11, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(11, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(11, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(11, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(11, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(11, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(11, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(11, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(11, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(11, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(11, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(11, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(11, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(11, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(11, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(11, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(11, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(11, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(11, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(11, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(11, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(11, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(11, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(11, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(11, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(11, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(11, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(11, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(11, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(11, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(11, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(11, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(11, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(11, 9).value = '3'
            Sheet1.range(11, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(11, 9).value = '4'
            Sheet1.range(11, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(11, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(11, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(11, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(11, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(11, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(11, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(11, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(11, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(11, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(11, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(11, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(11, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(11, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(11, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    #Final Exposure Results
    print(Final_MisinfoP1)
    print(Final_MisinfoP2)
    Sheet1.range(2, 19).value = len(Final_MisinfoP1)
    Sheet1.range(2, 20).value = len(Final_MisinfoP2)

    print('\n' + '~'*60 + '\n' + ' '*21 + 'ANALYSIS FINISHED PLEASE CHECK ERRORS AND EXCEL SHEET' + '\n' + '~'*60) # Aesthetic Break between Questions (Odd Spacing - can be one line)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 15 FINISHED
    # --------------------------------------------------------------------------------------------------------

#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
                # MORI
                # Discussion 4
                # Below
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------------------------

def moriD4():
    path = file_path.get()
    string = transcript_entry.get("1.0", "end")
    pair_number = pair_entry.get()
    total_text1 = string
    wb = xlwings.Book(path) #Finds Excel Book
    Sheet1 = wb.sheets[4] #grabs correct sheet (starts at zero - pythonic)
    
    Sheet1.range('A2','A3').value = pair_number #pair number
    Sheet1.range('A4','A5').value = pair_number
    Sheet1.range('A6','A7').value = pair_number
    Sheet1.range('A8','A9').value = pair_number
    Sheet1.range('A10','A11').value = pair_number
    
    Sheet1.range('B2','B3').value = 'F' #when the correct discussion is clicked, it auto populates the C/D/E/F
    Sheet1.range('B4','B5').value = 'F'
    Sheet1.range('B6','B7').value = 'F'
    Sheet1.range('B8','B9').value = 'F'
    Sheet1.range('B10','B11').value = 'F'
    
    total_text2 = total_text1.replace('!','.')
    total_text = total_text2.replace('?','.') #replace every question mark with a period to bypass regex

    valid = list('ABCDEFKMGT') + ['SAR','CAW','SDL','EWM','ERS']

    for w in re.findall('\w+', total_text):
        if w not in valid:
            total_text = re.sub(w, w.lower(), total_text)
    questions = total_text.split('~') #split each questions up into a list
    
    #Final Misinfo for each participant - across all questions for each discussion type
    Final_MisinfoP1 = []
    Final_MisinfoP2 = []

    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 1:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[0] #Which of the questions (list number) is bein read
    if 'question 1' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['A','B','D','C']

        Wrong_AnswerP2 = ['A','B','D','C']

        CWAB = ['A','B','D','C']
        
        dic = {'A':'SAR','B':'CAW','C':'SDL','D':'EWM','E':'ERS'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[1]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(2, 17).value = '1'
            Sheet1.range(2, 18).value = '2'
        elif 'person 2' in Chat[1]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(2, 17).value = '2'
            Sheet1.range(2, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'SAR' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'CAW' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'SDL' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'EWM' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'ERS' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
   

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'SAR' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'CAW' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'SDL' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'EWM' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'ERS' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(2, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(2, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(2, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(2, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(2, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(2, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(2, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(2, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')


        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(2, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(2, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(2, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(2, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(2, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(2, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(2, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(2, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(2, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(2, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(2, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(2, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(2, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(2, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(2, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(2, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(2, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(2, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(2, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(2, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(2, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 11).value = '2'
        
    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(2, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(2, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(2, 12).value = '2'
        
    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(2, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(2, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(2, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(2, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(2, 9).value = '3'
            Sheet1.range(2, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(2, 9).value = '4'
            Sheet1.range(2, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(2, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(2, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(2, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(2, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(2, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(2, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(2, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(2, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines
    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(2, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(2, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(2, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(2, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(2, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(2, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 1 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 2 IS A FILLER
    #------------------------

    #------------------------
    # QUESTION 3 IS A FILLER
    #------------------------


    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 4:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[3] #Which of the questions (list number) is bein read
    if 'question 4' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['A','B','D','C']

        Wrong_AnswerP2 = ['A','D','C','E']

        CWAB = ['A','C','D']
        
        dic = {'A':'F','B':'K','C':'M','D':'G','E':'T'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(8, 17).value = '1'
            Sheet1.range(8, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(8, 17).value = '2'
            Sheet1.range(8, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'F' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'K' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'M' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'G' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'T' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')


        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'F' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'K' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'M' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'G' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'T' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []
        check = any(item in CWAB for item in MisinfoP1)

        #Analysis
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(8, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(8, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(8, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(8, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(8, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(8, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(8, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(8, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(8, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(8, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(8, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(8, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(8, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(8, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(8, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(8, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(8, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(8, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(8, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(8, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(8, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(8, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(8, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(8, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(8, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(8, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(8, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(8, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(8, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(8, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(8, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(8, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(8, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(8, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(8, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(8, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(8, 9).value = '3'
            Sheet1.range(8, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(8, 9).value = '4'
            Sheet1.range(8, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(8, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(8, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(8, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(8, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(8, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(8, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(8, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(8, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(8, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(8, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(8, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(8, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(8, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(8, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 4 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 5:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[4] #Which of the questions (list number) is bein read
    if 'question 5' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'C'

        Wrong_AnswerP1 = ['E','B','D','A']

        Wrong_AnswerP2 = ['E','B','D','A']

        CWAB = ['E','B','D','A']
        
        dic = {'A':'no food or drink','B':'no smoking','C':'no mobile phones','D':'no cameras','E':'no animals'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(4, 17).value = '1'
            Sheet1.range(4, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(4, 17).value = '2'
            Sheet1.range(4, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'food' in each_sentence or 'person 2' in each_sentence and 'food or drink' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'smoking' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'mobile' in each_sentence or 'person 2' in each_sentence and 'phone' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'camera' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'animal' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
 

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'food' in each_sentence or 'person 1' in each_sentence and 'food or drink' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'smoking' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'mobile' in each_sentence or 'person 1' in each_sentence and 'phone' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'camera' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'animal' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(4, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(4, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(4, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(4, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(4, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(4, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(4, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(4, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(4, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(4, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(4, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(4, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(4, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(4, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(4, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(4, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(4, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(4, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(4, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(4, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(4, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(4, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(4, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(4, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(4, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(4, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(4, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(4, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(4, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(4, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(4, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(4, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(4, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(4, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(4, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(4, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(4, 9).value = '3'
            Sheet1.range(4, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(4, 9).value = '4'
            Sheet1.range(4, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(4, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(4, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(4, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(4, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(4, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(4, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(4, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(4, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(4, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(4, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(4, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(4, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(4, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(4, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 5 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 6:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[5] #Which of the questions (list number) is bein read
    if 'question 6' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'A'

        Wrong_AnswerP1 = ['D','E','B','A']

        Wrong_AnswerP2 = ['D','B','C','E']

        CWAB = ['B','D','E'] #combined wrong asnwers for both - used in exposure
        
        dic = {'A':'jug/pitcher','B':'wine glass','C':'mug','D':'soda can','E':'flask'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(10, 17).value = '1'
            Sheet1.range(10, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(10, 17).value = '2'
            Sheet1.range(10, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'jug' in each_sentence or 'person 2' in each_sentence and 'pitcher' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'wine' in each_sentence or 'person 2' in each_sentence and 'glass' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'mug' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'soda' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'flask' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
 

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'jug' in each_sentence or 'person 1' in each_sentence and 'pitcher' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'wine' in each_sentence or 'person 1' in each_sentence and 'glass' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'mug' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'soda' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'flask' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(10, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(10, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(10, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(10, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(10, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(10, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(10, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(10, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(10, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(10, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(10, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(10, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(10, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(10, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(10, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(10, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(10, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(10, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(10, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(10, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(10, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(10, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(10, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(10, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(10, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(10, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(10, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(10, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(10, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(10, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(10, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(10, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(10, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(10, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(10, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(10, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(10, 9).value = '3'
            Sheet1.range(10, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(10, 9).value = '4'
            Sheet1.range(10, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(10, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(10, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(10, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(10, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(10, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(10, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(10, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(10, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(10, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(10, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(10, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(10, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(10, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(10, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()


    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 6 FINISHED
    # --------------------------------------------------------------------------------------------------------


    #------------------------
    # QUESTION 7 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 8:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[7] #Which of the questions (list number) is bein read
    if 'question 8' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'A'

        Wrong_AnswerP1 = ['E','B','C','D']

        Wrong_AnswerP2 = ['E','B','C','D']

        CWAB = ['E','B','C','D']
        
        dic = {'A':'brown','B':'green','C':'grey','D':'purple','E':'black'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(3, 17).value = '1'
            Sheet1.range(3, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(3, 17).value = '2'
            Sheet1.range(3, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'brown' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'green' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'grey' in each_sentence or 'person 2' in each_sentence and 'gray' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'purple' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'black' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
     

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'brown' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'green' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'grey' in each_sentence or 'person 1' in each_sentence and 'gray' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'purple' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'black' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(3, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(3, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(3, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(3, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(3, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(3, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(3, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(3, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(3, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(3, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(3, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(3, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(3, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(3, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(3, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(3, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(3, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(3, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(3, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(3, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(3, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(3, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(3, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(3, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(3, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(3, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(3, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(3, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(3, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(3, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(3, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(3, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(3, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(3, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(3, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(3, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(3, 9).value = '3'
            Sheet1.range(3, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(3, 9).value = '4'
            Sheet1.range(3, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(3, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(3, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(3, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(3, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(3, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(3, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(3, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(3, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(3, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(3, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(3, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(3, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(3, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(3, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 8 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 9:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[8] #Which of the questions (list number) is bein read
    if 'question 9' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'D'
        Correct_AnswerP2 = 'A'

        Wrong_AnswerP1 = ['C','A','B','E']

        Wrong_AnswerP2 = ['D','B','E','C']

        CWAB = ['C','E','B']
        
        dic = {'A':'frying pan','B':'baking tray','C':'casserole dish','D':'sieve','E':'mixing bowl'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(7, 17).value = '1'
            Sheet1.range(7, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(7, 17).value = '2'
            Sheet1.range(7, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'frying pan' in each_sentence or 'person 2' in each_sentence and 'frying' in each_sentence or 'person 2' in each_sentence and 'pan' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'baking tray' in each_sentence or 'person 2' in each_sentence and 'tray' in each_sentence or 'person 2' in each_sentence and 'baking' in each_sentence or 'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'casserole dish' in each_sentence or 'person 2' in each_sentence and 'casserole' in each_sentence or 'person 2' in each_sentence and 'dish' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'sieve' in each_sentence or 'person 2' in each_sentence and 'strainer' in each_sentence or 'person 2' in each_sentence and 'sieve/strainer' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'mixing bowl' in each_sentence or 'person 2' in each_sentence and 'mixing' in each_sentence or 'person 2' in each_sentence and 'bowl' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
         

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'frying pan' in each_sentence or 'person 1' in each_sentence and 'frying' in each_sentence or 'person 1' in each_sentence and 'pan' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'baking tray' in each_sentence or 'person 1' in each_sentence and 'tray' in each_sentence or 'person 1' in each_sentence and 'baking' in each_sentence or 'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'casserole dish' in each_sentence or 'person 1' in each_sentence and 'casserole' in each_sentence or 'person 1' in each_sentence and 'dish' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'sieve' in each_sentence or 'person 1' in each_sentence and 'strainer' in each_sentence or 'person 1' in each_sentence and 'sieve/strainer' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'mixing bowl' in each_sentence or 'person 1' in each_sentence and 'mixing' in each_sentence or 'person 1' in each_sentence and 'bowl' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(7, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(7, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(7, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(7, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(7, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(7, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(7, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(7, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(7, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(7, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(7, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(7, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(7, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(7, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(7, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(7, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(7, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(7, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(7, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(7, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(7, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(7, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(7, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(7, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(7, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(7, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(7, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(7, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(7, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(7, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(7, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(7, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(7, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(7, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(7, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(7, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(7, 9).value = '3'
            Sheet1.range(7, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(7, 9).value = '4'
            Sheet1.range(7, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(7, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(7, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(7, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(7, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(7, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(7, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(7, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(7, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(7, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(7, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(7, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(7, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(7, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(7, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 9 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 10 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 11:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[10] #Which of the questions (list number) is bein read
    if 'question 11' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'E'
        Correct_AnswerP2 = 'E'

        Wrong_AnswerP1 = ['D','C','A','B']

        Wrong_AnswerP2 = ['D','C','A','B']

        CWAB = ['C','A','D','B']
        
        dic = {'A':'hard drive','B':'phone','C':'microphone','D':'tablet','E':'laptop'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(5, 17).value = '1'
            Sheet1.range(5, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(5, 17).value = '2'
            Sheet1.range(5, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'hard' in each_sentence or 'person 2' in each_sentence and 'drive' in each_sentence or 'person 2' in each_sentence and 'harddrive' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'phone' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'microphone' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'tablet' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'laptop' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
 

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'hard' in each_sentence or 'person 1' in each_sentence and 'drive' in each_sentence or 'person 1' in each_sentence and 'harddrive' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'phone' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'microphone' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'tablet' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'laptop' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(5, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(5, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(5, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(5, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(5, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(5, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(5, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(5, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(5, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(5, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(5, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(5, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(5, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(5, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(5, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(5, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(5, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(5, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(5, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(5, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(5, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(5, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(5, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(5, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(5, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(5, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(5, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(5, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(5, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(5, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(5, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(5, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(5, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(5, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(5, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(5, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(5, 9).value = '3'
            Sheet1.range(5, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(5, 9).value = '4'
            Sheet1.range(5, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(5, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(5, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(5, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(5, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(5, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(5, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(5, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(5, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(5, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(5, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(5, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(5, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(5, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(5, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 11 FINISHED
    # --------------------------------------------------------------------------------------------------------






    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 12:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[11] #Which of the questions (list number) is bein read
    if 'question 12' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'A'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['E','D','C','B']

        Wrong_AnswerP2 = ['A','C','E','D']

        CWAB = ['C','E','D']
        
        dic = {'A':'blue','B':'black','C':'pink','D':'white','E':'brown'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(9, 17).value = '1'
            Sheet1.range(9, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(9, 17).value = '2'
            Sheet1.range(9, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'blue' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'black' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'pink' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'white' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'brown' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')

            

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'blue' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'black' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'pink' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'white' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'brown' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(9, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(9, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(9, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(9, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(9, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(9, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(9, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(9, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(9, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(9, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(9, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(9, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(9, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(9, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(9, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(9, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(9, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(9, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(9, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(9, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(9, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(9, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(9, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(9, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(9, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(9, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(9, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(9, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(9, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(9, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(9, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(9, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(9, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(9, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(9, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(9, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(9, 9).value = '3'
            Sheet1.range(9, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(9, 9).value = '4'
            Sheet1.range(9, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(9, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(9, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(9, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(9, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(9, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(9, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(9, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(9, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(9, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(9, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(9, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(9, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(9, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(9, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 12 FINISHED
    # --------------------------------------------------------------------------------------------------------



    #------------------------
    # QUESTION 13 IS A FILLER
    #------------------------



    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 14:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[13] #Which of the questions (list number) is bein read
    if 'question 14' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'D'
        Correct_AnswerP2 = 'D'

        Wrong_AnswerP1 = ['A','E','C','B']

        Wrong_AnswerP2 = ['A','E','C','B']

        CWAB = ['A','C','E','B']
        
        dic = {'A':'blender','B':'kettle','C':'toaster','D':'microwave','E':'slow cooker'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(6, 17).value = '1'
            Sheet1.range(6, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(6, 17).value = '2'
            Sheet1.range(6, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'blender' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'kettle' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'toaster' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'microwave' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'slow' in each_sentence or 'person 2' in each_sentence and 'cooker' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
                

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'blender' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'kettle' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'toaster' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'microwave' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'slow' in each_sentence or 'person 1' in each_sentence and 'cooker' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')

        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(6, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(6, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(6, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(6, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(6, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(6, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(6, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(6, 13).value = '4'


    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(6, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(6, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(6, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(6, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(6, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(6, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(6, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(6, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(6, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(6, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(6, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(6, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(6, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(6, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(6, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(6, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(6, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(6, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(6, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(6, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(6, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(6, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(6, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(6, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(6, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(6, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(6, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(6, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(6, 9).value = '3'
            Sheet1.range(6, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(6, 9).value = '4'
            Sheet1.range(6, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(6, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(6, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(6, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(6, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(6, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(6, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(6, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(6, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(6, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(6, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(6, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(6, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(6, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(6, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 14 FINISHED
    # --------------------------------------------------------------------------------------------------------







    #-------------------------------------------------------------------------------------------------------------------
    #BEGINNING OF QUESTION 15:
    #-------------------------------------------------------------------------------------------------------------------

    txt = questions[14] #Which of the questions (list number) is bein read
    if 'question 15' in txt:  # Makes sure you are inputting the correct question
        Chat_with_experimenter = txt.split('\n')  # Split the text into sentences divided by a period
        Final_AnswerP1 = ''  # Final Answer is filled depending on what it is
        Final_AnswerP2 = ''  
        MisinfoP1 = [] 
        MisinfoP2 = [] # Open to be filled by each if-statement about what specific misinfo occured. Used later.

        Correct_AnswerP1 = 'C'
        Correct_AnswerP2 = 'B'

        Wrong_AnswerP1 = ['A','D','B','E']

        Wrong_AnswerP2 = ['A','E','C','D']

        CWAB = ['A','E','D']
        
        dic = {'A':'nothing, the screen is off','B':'abstract pattern','C':'car advert','D':'news channel','E':'timetable'}

        print('\n' + '-'*60 + '\n' + ' '*21 + txt[0:12] + '\n' + '-'*60 + '\n')# Aesthetic Break between Questions

        Chat = [sentence for sentence in Chat_with_experimenter if 'experimenter' not in sentence] #deletes all sentences that have 'experimenter' in it.

        print("\u0332".join('\n\nSPEAKER ORDER:\n'))  # This underlines

        if 'person 1' in Chat[2]: #This determines who spoke first
            print('Person 1 spoke first')
            print('Person 2 spoke second')
            Sheet1.range(11, 17).value = '1'
            Sheet1.range(11, 18).value = '2'
        elif 'person 2' in Chat[2]:
            print('Person 2 spoke first')
            print('Person 1 spoke second')
            Sheet1.range(11, 17).value = '2'
            Sheet1.range(11, 18).value = '1'
        else:
            print('error')  

    # Misinformation Loop for Person 1

        print("\u0332".join('\n\nMISINFORMATION INFO:\n'))  # This underlines 'Misinfo'
        for each_sentence in Chat:
            if 'person 2' in each_sentence and 'nothing' in each_sentence or 'person 2' in each_sentence and 'off.' in each_sentence or 'person 2' in each_sentence and 'off ' in each_sentence or 'person 2' in each_sentence and 'A' in each_sentence:
                MisinfoP1.append('A')
                 #Unused but here for later uses
            if 'person 2' in each_sentence and 'abstract' in each_sentence or 'person 2' in each_sentence and 'pattern' in each_sentence or'person 2' in each_sentence and 'B' in each_sentence:
                MisinfoP1.append('B')
                            
            if 'person 2' in each_sentence and 'car' in each_sentence or 'person 2' in each_sentence and 'advert' in each_sentence or 'person 2' in each_sentence and 'C' in each_sentence:
                MisinfoP1.append('C')
                            
            if 'person 2' in each_sentence and 'news' in each_sentence or 'person 2' in each_sentence and 'channel' in each_sentence or 'person 2' in each_sentence and 'D' in each_sentence:
                MisinfoP1.append('D')
                            
            if 'person 2' in each_sentence and 'time' in each_sentence or 'person 2' in each_sentence and 'E' in each_sentence:
                MisinfoP1.append('E')
         

        FinalMisinfoP1 = [x for x in MisinfoP1 if Correct_AnswerP1 not in x] #Gets rid of any experienced values as a new variable so that I can use MisinfoP1 later on

        print(f'Person 1 was exposed to the following misinformation: {FinalMisinfoP1}')

    # Misinformation Loop for Person 2

        for each_sentence in Chat:
            if 'person 1' in each_sentence and 'nothing' in each_sentence or 'person 1' in each_sentence and 'off.' in each_sentence or 'person 2' in each_sentence and 'off ' in each_sentence or 'person 1' in each_sentence and 'A' in each_sentence:
                MisinfoP2.append('A')
                 #Unused but here for later uses
            if 'person 1' in each_sentence and 'abstract' in each_sentence or 'person 1' in each_sentence and 'pattern' in each_sentence or'person 1' in each_sentence and 'B' in each_sentence:
                MisinfoP2.append('B')
                            
            if 'person 1' in each_sentence and 'car' in each_sentence or 'person 1' in each_sentence and 'advert' in each_sentence or 'person 1' in each_sentence and 'C' in each_sentence:
                MisinfoP2.append('C')
                            
            if 'person 1' in each_sentence and 'news' in each_sentence or 'person 1' in each_sentence and 'channel' in each_sentence or 'person 1' in each_sentence and 'D' in each_sentence:
                MisinfoP2.append('D')
                            
            if 'person 1' in each_sentence and 'time' in each_sentence or 'person 1' in each_sentence and 'E' in each_sentence:
                MisinfoP2.append('E')
                

        FinalMisinfoP2 = [x for x in MisinfoP2 if Correct_AnswerP2 not in x] #Gets rid of any experienced values

        print(f'Person 2 was exposed to the following misinformation: {FinalMisinfoP2}')

    #EXPOSURE INFORMATION - This determines the exposure type and DOES NOT remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 1 INFO:\n'))  # This underlines

        exposureP1 = []

        check = any(item in CWAB for item in MisinfoP1)

        if Correct_AnswerP2 in MisinfoP1:
            exposureP1.append('1')
        if Correct_AnswerP1 in MisinfoP1:
            exposureP1.append('3')
        if check is True:
            exposureP1.append('2')
        if MisinfoP1 == []:
            exposureP1.append('0')


        #Computation
        if exposureP1 == ['0'] or exposureP1 == []:
            print('0')
            Sheet1.range(11, 13).value = '0'
        if exposureP1 == ['1']:
            print('1')
            Sheet1.range(11, 13).value = '1'
        if exposureP1 == ['2']:
            print('2')
            Sheet1.range(11, 13).value = '2'
        if exposureP1 == ['3']:
            print('3')
            Sheet1.range(11, 13).value = '3'

        if '1' in exposureP1 and '2' in exposureP1 and '3' in exposureP1:
            print('7')
            Sheet1.range(11, 13).value = '7'
        elif '2' in exposureP1 and '3' in exposureP1:
            print('6')
            Sheet1.range(11, 13).value = '6'
        elif '1' in exposureP1 and '3' in exposureP1:
            print('5')
            Sheet1.range(11, 13).value = '5'
        elif '1' in exposureP1 and '2' in exposureP1:
            print('4')
            Sheet1.range(11, 13).value = '4'

    #EXPOSURE INFORMATION - This determines the exposure type and does not remove any experienced values

        print("\u0332".join('\n\nEXPOSURE PERSON 2 INFO:\n'))  # This underlines

        exposureP2 = []

        check = any(item in CWAB for item in MisinfoP2)

        if Correct_AnswerP2 in MisinfoP2:
            exposureP2.append('3')
        if Correct_AnswerP1 in MisinfoP2:
            exposureP2.append('1')
        if check is True:
            exposureP2.append('2')
        if MisinfoP2 == []:
            exposureP2.append('0')

        #Computation
        if exposureP2 == ['0'] or exposureP2 == []:
            print('0')
            Sheet1.range(11, 14).value = '0'
        if exposureP2 == ['1']:
            print('1')
            Sheet1.range(11, 14).value = '1'
        if exposureP2 == ['2']:
            print('2')
            Sheet1.range(11, 14).value = '2'
        if exposureP2 == ['3']:
            print('3')
            Sheet1.range(11, 14).value = '3'

        if '1' in exposureP2 and '2' in exposureP2 and '3' in exposureP2:
            print('7')
            Sheet1.range(11, 14).value = '7'
        elif '2' in exposureP2 and '3' in exposureP2:
            print('6')
            Sheet1.range(11, 14).value = '6'
        elif '1' in exposureP2 and '3' in exposureP2:
            print('5')
            Sheet1.range(11, 14).value = '5'
        elif '1' in exposureP2 and '2' in exposureP2:
            print('4')
            Sheet1.range(11, 14).value = '4'

    # Final Answer Loop for Person 1

        print()
        print("\u0332".join('FINAL ANSWER PERSON 1 INFO: \n'))  # Underlined
        if 'p1 F' in Chat[-3] and 'A' in Chat[-3]: #Uses -3 because the final period in the text file creates an extra list item in Chat
            print('Person 1 Final Answer is A')
            Final_AnswerP1 = 'A'
            Sheet1.range(11, 7).value = dic['A']
        elif 'p1 F' in Chat[-3] and 'B' in Chat[-3]:
            print('Person 1 Final Answer is B')
            Final_AnswerP1 = 'B'
            Sheet1.range(11, 7).value = dic['B']
        elif 'p1 F' in Chat[-3] and 'C' in Chat[-3]:
            print('Person 1 Final Answer is C')
            Final_AnswerP1 = 'C'
            Sheet1.range(11, 7).value = dic['C']
        elif 'p1 F' in Chat[-3] and 'D' in Chat[-3]:
            print('Person 1 Final Answer is D')
            Final_AnswerP1 = 'D'
            Sheet1.range(11, 7).value = dic['D']
        elif 'p1 F' in Chat[-3] and 'E' in Chat[-3]:
            print('Person 1 Final Answer is E')
            Final_AnswerP1 = 'E'
            Sheet1.range(11, 7).value = dic['E']
        else:
            print('Error - final answer for Person 1 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP1 = ''
            Sheet1.range(11, 7).value = 'ERROR'

    # Final Answer Loop for Person 2

        print()
        print("\u0332".join('FINAL ANSWER PERSON 2 INFO: \n'))  # Underlined
        if 'p2 F' in Chat[-2] and 'A' in Chat[-2]: #Uses -2 because the final period in the text file creates an extra list item in Chat
            print('Person 2 Final Answer is A')
            Final_AnswerP2 = 'A'
            Sheet1.range(11, 8).value = dic['A']
        elif 'p2 F' in Chat[-2] and 'B' in Chat[-2]:
            print('Person 2 Final Answer is B')
            Final_AnswerP2 = 'B'
            Sheet1.range(11, 8).value = dic['B']
        elif 'p2 F' in Chat[-2] and 'C' in Chat[-2]:
            print('Person 2 Final Answer is C')
            Final_AnswerP2 = 'C'
            Sheet1.range(11, 8).value = dic['C']
        elif 'p2 F' in Chat[-2] and 'D' in Chat[-2]:
            print('Person 2 Final Answer is D')
            Final_AnswerP2 = 'D'
            Sheet1.range(11, 8).value = dic['D']
        elif 'p2 F' in Chat[-2] and 'E' in Chat[-2]:
            print('Person 2 Final Answer is E')
            Final_AnswerP2 = 'E'
            Sheet1.range(11, 8).value = dic['E']
        else:
            print('Error - final answer for Person 2 must be A, B, C, D, or E. Please check your input')
            Final_AnswerP2 = ''
            Sheet1.range(11, 8).value = 'ERROR'

    # Misled Loop for Person 1

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 1: \n'))  # Underlined

        if Final_AnswerP1 == Correct_AnswerP1:
            print(f'Person 1 was not misled because they answered correctly with: {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '2'
        elif Final_AnswerP1 in MisinfoP1:#Don't have to use FinalMisinfo because if the final answer is exposed, then it won't reach this statement. 
            print(f'Person 1 was misled to answer {Final_AnswerP1}')
            Sheet1.range(11, 11).value = '1'
        else:
            print(f'Person 1 was not misled because, although they answered incorrectly - {Final_AnswerP1} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 11).value = '2'

    # Misled Loop for Person 2

        print()
        print("\u0332".join('FINAL MISLED INFO PERSON 2: \n'))  # Underlined

        if Final_AnswerP2 == Correct_AnswerP2:
            print('Person 2 was not misled because they answered correctly with: ' + Final_AnswerP2)
            Sheet1.range(11, 12).value = '2'
        elif Final_AnswerP2 in MisinfoP2:
            print(f'Person 1 was misled to answer {Final_AnswerP2}')
            Sheet1.range(11, 12).value = '1'
        else:
            print(f'Person 2 was not misled because, although they answered incorrectly - {Final_AnswerP2} - it was not previously exposed to them as misinfo')
            Sheet1.range(11, 12).value = '2'

    # Discussion Type Loop

        print("\u0332".join('\n\nDISCUSSION TYPE INFO:\n'))  # This underlines

                    #AGREED

        #Pair agreed on participant's witnessed detail
        #Pair agreed on participant's witnessed detail
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP1 == Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1: Code 1')
            Sheet1.range(11, 9).value = '1'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP2 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 2: Code 1')
            Sheet1.range(11, 10).value = '1'

        #Pair agreed on partner's witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 and Correct_AnswerP2 == Final_AnswerP1 and Final_AnswerP1 != Correct_AnswerP1:
            print('Person 1: Code 2')
            Sheet1.range(11, 9).value = '2'
        if Final_AnswerP2 == Final_AnswerP1 and Correct_AnswerP1 == Final_AnswerP2 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 2: Code 2')
            Sheet1.range(11, 10).value = '2'

        #Pair agreed on joint witnessed detail:
        if Final_AnswerP1 == Final_AnswerP2 == Correct_AnswerP1 == Correct_AnswerP2:
            print('Person 1 and Person 2: Code 3')
            Sheet1.range(11, 9).value = '3'
            Sheet1.range(11, 10).value = '3'

        #Pair agreed on other inaccurate detail:
        if Final_AnswerP1 == Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2:
            print('Person 1 and Person 2: Code 4')
            Sheet1.range(11, 9).value = '4'
            Sheet1.range(11, 10).value = '4'

                  #DISAGREED (first two could be mirrors....?)

        #Participant provided participant-witnessed detail
        if Final_AnswerP1 == Correct_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5a')
            Sheet1.range(11, 9).value = '5a'
        if Final_AnswerP2 == Correct_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5a')
            Sheet1.range(11, 10).value = '5a'

        #Participant provided partner-witnessed detail
        if Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 == Correct_AnswerP2 and Final_AnswerP1 != Final_AnswerP2:
            print('Person 1: Code 5b')
            Sheet1.range(11, 9).value = '5b'
        if Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 == Correct_AnswerP1 and Final_AnswerP2 != Final_AnswerP1:
            print('Person 2: Code 5b')
            Sheet1.range(11, 10).value = '5b'

        #Participant provided other inaccurate detail mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 in MisinfoP1:
            print('Person 1: Code 5c')
            Sheet1.range(11, 9).value = '5c'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 in MisinfoP2:
            print('Person 2: Code 5c')
            Sheet1.range(11, 10).value = '5c'

        #Participant provided other inaccurate detail NOT mentioned by partner
        if Final_AnswerP1 != Final_AnswerP2 and Final_AnswerP1 != Correct_AnswerP1 and Final_AnswerP1 != Correct_AnswerP2 and Final_AnswerP1 not in MisinfoP1:
            print('Person 1: Code 5d')
            Sheet1.range(11, 9).value = '5d'
        if Final_AnswerP2 != Final_AnswerP1 and Final_AnswerP2 != Correct_AnswerP2 and Final_AnswerP2 != Correct_AnswerP1 and Final_AnswerP2 not in MisinfoP2:
            print('Person 2: Code 5d')
            Sheet1.range(11, 10).value = '5d'
        #Codes 5e and 5f are Not Applicable! 

    #CONFIDENCE

        print("\u0332".join('\n\nCONFIDENCE:\n'))  # This underlines

    #Person 1 Confidence    
        if 'p1 conf' in Chat[-5] and 'l' in Chat[-5]:
            print('Person 1 Confidence is 1')
            Sheet1.range(11, 15).value = '1'
        if 'p1 conf' in Chat[-5] and 'm' in Chat[-5]:
            print('Person 1 Confidence is 2')
            Sheet1.range(11, 15).value = '2'
        if 'p1 conf' in Chat[-5] and 'h' in Chat[-5]:
            print('Person 1 Confidence is 3')
            Sheet1.range(11, 15).value = '3'

    #Person 2 Confidence
        if 'p2 conf' in Chat[-4] and 'l' in Chat[-4]:
            print('Person 2 Confidence is 1')
            Sheet1.range(11, 16).value = '1'
        if 'p2 conf' in Chat[-4] and 'm' in Chat[-4]:
            print('Person 2 Confidence is 2')
            Sheet1.range(11, 16).value = '2'
        if 'p2 conf' in Chat[-4] and 'h' in Chat[-4]:
            print('Person 2 Confidence is 3')
            Sheet1.range(11, 16).value = '3'
            
    #Final Exposure Information
    Final_MisinfoP1.extend(MisinfoP1)
    Final_MisinfoP2.extend(MisinfoP2)
    
    #Final Exposure Results
    print(Final_MisinfoP1)
    print(Final_MisinfoP2)
    Sheet1.range(2, 19).value = len(Final_MisinfoP1)
    Sheet1.range(2, 20).value = len(Final_MisinfoP2)

    print('\n' + '~'*60 + '\n' + ' '*21 + 'ANALYSIS FINISHED PLEASE CHECK ERRORS AND EXCEL SHEET' + '\n' + '~'*60) # Aesthetic Break between Questions (Odd Spacing - can be one line)

    wb.save()

    # --------------------------------------------------------------------------------------------------------
    # Block Code Separation Line - QUESTION 15 FINISHED
    # --------------------------------------------------------------------------------------------------------


#################################################
#           GRAPHICAL USER INTERFACE            # 
#################################################

root = tkinter.Tk()
root.title('MORI Transcript Auto-Coder') #Title of Program

canvas1 = tkinter.Canvas(root, height=650, width=1800)
canvas1.grid(columnspan = 3, rowspan=3)
canvas1.configure(background='gray25')
root.geometry("1000x650")

title = tkinter.Label(root, text='MORI', bg = 'gray25', fg='black')
title.config(font=('Times', 50, 'bold italic'))
canvas1.create_window(500, 50, window=title)

subtitle = tkinter.Label(root, text='Transcript Auto-Coder', bg = 'gray25', fg='black')
subtitle.config(font=('Times', 15, 'bold italic'))
canvas1.create_window(495, 90, window=subtitle)

transcript_title = tkinter.Label(root, text='Transcript', bg = 'gray25', fg='white') #Label for e
transcript_title.config(font=('helvetica', 20, "bold"))
canvas1.create_window(500, 150, window=transcript_title)

file_path_title = tkinter.Label(root, text='File Path', bg = 'gray25', fg='white') #Label for h
file_path_title.config(font=('helvetica', 10, "bold"))
canvas1.create_window(925, 150, window=file_path_title)

pair_number_title = tkinter.Label(root, text='Pair Number', bg = 'gray25', fg='white') #Label for p
pair_number_title.config(font=('helvetica', 10, "bold"))
canvas1.create_window(75, 150, window=pair_number_title)

file_path = tkinter.Entry (root)
file_path.insert(0, r"C:\Users")
canvas1.create_window(925, 175, window=file_path)

transcript_entry = tkinter.Text (root) 
canvas1.create_window(500, 360, window=transcript_entry) #entry for label 1

pair_entry = tkinter.Entry (root)
canvas1.create_window(75, 175, window=pair_entry)

b1 = tkinter.Button(root, text='Run Discussion 1 / C', command=moriD1, bg='brown3', fg='black', font=('helvetica', 10, 'bold'))
canvas1.create_window(125, 625, window=b1)


b2 = tkinter.Button(root, text='Run Discussion 2 / D', command=moriD2, bg='sienna4', fg='black', font=('helvetica', 10, 'bold'))
canvas1.create_window(375, 625, window=b2)


b3 = tkinter.Button(root, text='Run Discussion 3 / E', command=moriD3, bg='forestgreen', fg='black', font=('helvetica', 10, 'bold'))
canvas1.create_window(625, 625, window=b3)


b4 = tkinter.Button(root, text='Run Discussion 4 / F', command=moriD4, bg='SkyBlue1', fg='black', font=('helvetica', 10, 'bold'))
canvas1.create_window(875, 625, window=b4)

def delete1():
    file_path.delete(0,'end')
def delete2():
    transcript_entry.delete('1.0','end')
def delete3():
    pair_entry.delete(0,'end')
 
delete_1 = tkinter.Button(root, text = "Delete Field", command = delete1, bg = 'white', fg='black') # h delete
canvas1.create_window(925, 225, window=delete_1)

delete_2 = tkinter.Button(root, text = "Delete Field", command = delete2, bg = 'white', fg='black') # e delete
canvas1.create_window(500, 575, window=delete_2)

delete_3 = tkinter.Button(root, text = "Delete Field", command = delete3, bg = 'white', fg='black') # p delete
canvas1.create_window(75, 225, window=delete_3)




diagnostics = tkinter.Label(root, text='Diagnostics', bg = 'gray25', fg='red')
diagnostics.config(font=('Times', 15, 'bold italic'))
canvas1.create_window(1400, 150, window=diagnostics)
textbox=Text(root)
canvas1.create_window(1400, 365, window=textbox)
def redirector(inputStr):
    textbox.insert(INSERT, inputStr)
sys.stdout.write = redirector #whenever sys.stdout.write is called, redirector is called.

root.mainloop()

# In[ ]:




