#Last Updated 2nd July 2018
#Written by Kevin Wang and Bryan Bollou
#Contact Information
#Email: kevwang2292@gmail.com
#linkedin.com/in/kevwang2292


#import requests and beautifulsoup code libraries to use their functions
import requests
from bs4 import BeautifulSoup

# importing python libraries for reading from xlsx files (openpyxl) and interacting in the directory (os)
import os
from openpyxl import Workbook

#Get the target filename to create
file = input("What is filename you want to write to? (don't have to add the extension): ")

#Create excel workbook then set it as the active stream
wb = Workbook()
active = wb.active

#Create the title columns in the excel sheet
active ["A1"] = "Term"
active ["B1"] = "Definition"

#set the counter for the position in the excel sheet
counter2 = 1

#define function to parse a target page 
def parsepage(page, counter2):   

    #request the html object of the target url
    request = requests.get(page)

    #convert the returned html object into a beautifulsoup object
    soup = BeautifulSoup(request.content, "html.parser")

    #convert the beautifulsoup object into a string
    content = soup.prettify()

	#set counter for the position in the content
    counter = 0
    #create the variable for the next starting point of the term
    nextstart = 0

	#while loop to continously go through the content parsed from the target url
    while (True):
	
	    #find the starting and ending term positions starting from the last found term (0 if it is the first run through)
        startt = content.find("<h3>", counter)
        endt = content.find("</h3>", counter)

		#find where the next term starts
        nextstart = content.find("<h3>", endt+4)
	
		#find where the starting and ending definition positing are starting from the end of a found term
        startd = content.find("<p>", endt)
        endd = content.find("</p>", startd)
	
		#if a definition is not found with <p> or it is found to be for the next term then find it with <br/>
        if startd == -1 or startd > nextstart:
            startd = content.find("<br/>", endt)
            endd = content.find("<br/>", startd+4)
		
		#if a term is found, write it and the definition to the excel sheet
        if startd != -1:
			#term is between the starting and ending term positions
            term = (content [startt+4:endt]).strip()

			#definition is between the starting and ending definition positions
            definition = (content [startd+5:endd]).strip()
			
			#defining all exceptions for the first <p> definition
            badtags = ["<em>", "<ol>", "<li>", "<strong>", "<span", "<h3>"]
			
			#checking for bag tags in the definition
            for tag in badtags:
			    #Exception for definitions that contain titles such as "Telephony Term" - start the definition from the second <p> which is after the first </p> and after the </em>
                if tag in str(definition):
			        #find the end of the title term and set it a variable called x
                    x = content.find("</p>", endd)
				    #if a following term title exists
                    if content.find("<p>", x) != -1:
					    #find where the term title starts
                        startd = content.find("<p>", x)
                        if startd != -1:
					        #find where the term title ends
                            endd2 = content.find("</p>", startd)
				
				    #recreate the definition of the exception
                    definition = (content [startd+5:endd2]).strip()
                    
            #remove certain tags from definition
            for tag in ["<p>", "<em>", "<strong>", "<ul>", "<il>", "<span>", "</span>", "</p>", "</em>", "</strong>", "</ul>", "</il>"]:
                if tag in str(definition):
                    print("ran")
                    definition = (definition.replace(tag, "")).strip()     
							
		    #set counter to the end of the found definition to look for the next term from here
            counter = endd + 4
		
			#move down a row in the excel sheet
            counter2 += 1
		
			#write the term in the A column in the appropriate row
            coordinate = 'A' + str(counter2)
            active [coordinate] = term
			
		    #write the definition in the B column in the appropriate row
            coordinate = 'B' + str(counter2)
            active [coordinate] = definition
		
		#if no term is found
        else:
		    #return last written row number in the excel sheet
            return counter2
			#break out of the while loop to start on the next page or end the program
            break;

#Ask the user for the name of text file with the list of links
file2 = input("What is filename of the text file with the list of links? (don't have to add the extension): ")
			
#open file reader to read the list of links from a text file
filelist = open(str(file2)+".txt", "r")

#set the lines read into a list called "linklist"
linklist = filelist.readlines()

#iterate through the list and run the parsepage function for each list while starting from the ending position of the last run with counter2
for link in linklist:
    counter2 = (parsepage(link, counter2))
	    
#save the excel sheet with all of its new changes
wb.save(file+".xlsx")

    