import xlrd

wb = xlrd.open_workbook("Sem7 TimeTable - 3.xlsx")
sheet = wb.sheet_by_index(0) 


#3,1 - #14,4

def presentInList(li,item):
    for l in li:
        if item == l:
            return True
    return False

def containsDash(s):
    for i in range(len(s)):
        if(s[i] == "-"):
            return i
    return -1

numSubjects = int(input("Enter number of subjects - "))

subList = []
print("Enter the subject code in CAPS - ")
for i in range(numSubjects):
    sub = input();
    subList.append(sub)

#print(subList)

batchCode = input("Enter your batch code - ")


TimeAt = {1:"9:00", 2:"10:00", 3:"11:00", 4:"12:00",6:"2:00",7:"3:00",8:"4:00"}

dayRangeI = [3,17,27,39,49,60]
dayRangeJ = [15,26,38,48,59,68]

dayIs = {0:"Monday", 1:"Tuesday", 2:"Wednesday", 3:"Thursday", 4:"Friday", 5:"Saturday"}

subDict = { "NEC735":"Information Theory and Applications", "NEC736": "Essentials of VLSI Testing",  "EC412" : "Multimedia Communications",  "EC733":"Optical Communications",
                "EC416": "Deep Learning for Multimedia" , "NEC733":"Fundamental of Embedded Systems"  ,  "NEC734":"RF and Microwave Engg."  ,"EC731":"Mobile Communication"  ,
            "EC732":"Cognitive Communication Systems"  ,"MA831":"Applied Numerical Methods"  ,"MA731":"Applied Linear Algebra"  ,"MA411":"Elements Of Statistical Learning"  ,
            "PH732":"Nano Science And Technology"  ,"HS731":"Customer Relationship Management"  ,"HS732":"Indian Financial System"  ,"HS831":"Gender Studies"  ,
            "HS412":"Human Resource Analytics"  ,"HS733":"Human Rights And Social Justice"  ,"CI833":"Nature Inspired Computing"  ,"CS437":"Large Scale Database Systems"  ,
            "17B1NCI731":"Machine Learning and NLP"  ,"15B1NCI732":"Social Network Analysis"  ,"17B1NCI732":"Computer and Web security"  ,"CI747":"Cloud Computing"  ,
            "CS422":"Mathematical Foundations for Intelligent Systems"  ,"CS434":"Ethical Hacking"  ,"CS436":"Software Construction"  ,"CS423":"Computing for data science"  ,
            "CS731":"Data Compression Algorithms"  ,"CS424":"Industrial Automation using IoT"  ,"CS426":"IoT analytics"  ,"CS427":"Introduction to DEVOPS"  ,
            "17B2NCI731":"Computer Graphics"} 

for day in range(6):
    print()
    
    print(dayIs[day])
    for i in range(1,9):
        
        if i == 5:
            continue
        
        for j in range(dayRangeI[day],dayRangeJ[day]):
            
            cell = sheet.cell_value(j, i)
            if cell != "":
                lecType = cell[0]
                k = 1
                batchesInThisCell = ""
                batchList = []
                #print(j," ",i)
                while(cell[k] != '(' ):
                    if(cell[k] == ","):
                        batchList.append(batchesInThisCell)
                        batchesInThisCell = ""
                    else:
                        batchesInThisCell = batchesInThisCell + cell[k]
                    k = k+1
                batchList.append(batchesInThisCell)
                subjectCodeInThisCell = ""
                k = k+1
                while(cell[k] != ')' ):
                    subjectCodeInThisCell = subjectCodeInThisCell + cell[k]
                    k = k+1

                if presentInList(subList,subjectCodeInThisCell):
                
                    expandedBatchList = []
                    
                    for b in batchList:
                        bCode = b[0]
                        if bCode == "M":
                            bCode = "MBT"
                            bList = b[3:]
                        else:
                            bList = b[1:]
                        if(bList == ""):
                            if bCode == "A":
                                rangeStart = 1
                                rangeEnd = 4
                            elif bCode == "B":
                                rangeStart = 1
                                rangeEnd = 15
                            elif bCode == "C":
                                rangeStart = 1
                                rangeEnd = 2
                            elif bCode == "MBT":
                                rangeStart = 1
                                rangeEnd = 2
                            for idx in range(rangeStart,rangeEnd+1):
                                expandedBatchList.append(bCode+str(idx))
                                
                        #print(b," ",bCode," ",bList,end=" ")
                        dashAt = containsDash(bList)
                        if  dashAt != -1:
                            rangeStart = int(bList[0:dashAt])
                            rangeEnd = int(bList[dashAt+1:])
                            #print(rangeStart," ",rangeEnd)
                            for idx in range(rangeStart,rangeEnd+1):
                                expandedBatchList.append(bCode+str(idx))
                        else:
                            if bList != "":
                                expandedBatchList.append(bCode+bList)

                    if(presentInList(expandedBatchList,batchCode)):
                        k = k+2
                        classRoom = ""
                        while(cell[k] != "/"):
                            classRoom = classRoom + cell[k]
                            k = k+1
                        print(TimeAt[i],end = "   ")
                        classType = ""
                        if lecType == "L":
                            classType = "Lecture"
                        elif lecType == "T":
                            classType = "Tutorial"
                        print(classType,"- ",subDict[subjectCodeInThisCell], " Room - ",classRoom)

                
                    
                    
                
            
                               
