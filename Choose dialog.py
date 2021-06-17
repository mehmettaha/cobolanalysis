#!/usr/bin/env python
# coding: utf-8

# In[3]:


import os
import pandas as pd
from re import split as SPLIT
import sys
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox


# In[4]:


class Cobolfile:
 
    def __init__(self,name):
        self.kanren_files = set()
        self.program_id = ""
        self.revise_date = set()
        self.calllist  = set()
        self.sources = "*NoSource*"
        self.copylist = set()
        self.filename = name.replace(" ","_")
        self.author = "*NoAuthor*"
        self.kanren_programs = set()
        self.recorddict = dict()
        self.write_files = set()
        self.copy_files = set()
    
    def prep(self):
        print("File name:" + self.filename)
        print("Related files:" + str(self.kanren_files))
        print("Dates revised:" + str(self.revise_date))
        print("Author:" + str(self.author))
        print("Program ID:" + str(self.program_id))
        print("Source machine:" + str(self.sources))
        print("Record names for files:" + str(self.recorddict))
        print("Files written:" + str(self.write_files))
        print("Files copied:" + str(self.copy_files))
        print("Programs called:" + str(self.kanren_programs))

    def find_write_file(self, recname):   #Find and replace the inner file names for record names
        for i in self.recorddict.items():
            if i[1] == recname:
                return i[0]
        return None
        
class Jobfile:
    
    def __init__(self,name):
        self.filename = name.replace(" ","_")
        self.job_id = ""
        self.copy_files = set()
        self.assign = set()
        self.use_pairs = set()
        self.go_to_main = set()
        self.xqt = set()
        
    def prep(self):
        print("File name is:" + self.filename)
        print("Job ID is:" + self.job_id)
        print("Copied files are:" + str(self.copy_files))
        print("Assigned files are:" + str(self.assign))
        print("File names used:" + str(self.go_to_main))
        print("Executed programs:" + str(self.xqt))
  
    def recur(self, first, pairlist,collect, naibulist, new = "*"):   #Find real file names for inner file names used in job file
        if new == "*":
            new = first
        flag = 0
        for line in pairlist:
            if line[0] == new:
                xx = line[1]
                flag = 1
                self.recur(first, pairlist, collect,naibulist, new = xx)

        if flag == 0:        
            collect.add((first, new))
            naibulist.add(first)
            
        


# In[5]:


class Excelize:
    #Main Tables. Initial names are set here, but the graph tool still supports customizations. 
    jobcheck = 0       #Check if joblist is prepared for excel
    programcheck = 0   #Check if cobollist is prepared for excel
    classtable = pd.DataFrame([[1,"ProgramFile","rectangle","","aquamarine",""],        
                               [2,"RevisionHistory","rectangle","","lightblue",""],
                               [3,"Author","rectangle","blue","gray99",""],
                               [4,"ProgramID","rectangle","","lightgreen",""],
                               [5,"SourceMachine","rectangle","green","gray99",""],
                               [6, "CopiedFiles","rectangle","","pink",""],
                               [7, "NaibuFiles", "rectangle","","gray",""],
                               [8 , "Assign", "rectangle", "", "yellow", ""],
                               [9, "JobFile", "rectangle","","gray99",""],
                               [10, "JobID", "rectangle","","beige",""]],
                               columns = ["No", "Class Name", "Shape", "Line Color", "Fill Color", "remarks"]).set_index("No")
    relationtable = pd.DataFrame([[1, "ProgramFile", "ProgramID"  ,"", "blueviolet", "relation1", "", ""],
                                  [2, "Author","ProgramID", "", "brown3","relation2", "", ""],
                                  [3, "SourceMachine", "ProgramID", "","cadetblue4","relation3", "", ""],
                                  [4, "ProgramID", "CopiedFiles", "", "red", "relation4", "copy", ""],
                                  [5, "ProgramID", "ProgramID", "", "blue", "relation5", "call", ""],
                                  [6, "RevisionHistory","ProgramFile", "", "","relation6", "", "back"],
                                  [7, "ProgramID", "NaibuFiles", "bold","yellow", "relation7", "write", ""],
                                  [8, "ProgramID", "NaibuFiles", "","","relation8","",""],
                                  [9 , "JobFile", "JobID", "", "", "relation9", "",""],
                                  [10, "JobID", "Assign", "","","relation10", "Assign",""],
                                  [11, "JobID","ProgramID","dashed", "darkorange", "relation11", "execute", ""]],
                                   columns = ["No","Class Name 1","Class Name 2", "Line Type", "Line Color", "remarks", "Relation Type", "Direction"]).set_index("No")
   #Instance Tables
    header_list = ["No", "Instance Name", "Shape", "Line Color", "Fill Color", "remarks"]
    Instance_ProgramFile = pd.DataFrame(columns = header_list).set_index("No")
    Instance_RevisionHistory = pd.DataFrame(columns = header_list).set_index("No")
    Instance_Author = pd.DataFrame(columns = header_list).set_index("No")
    Instance_ProgramID = pd.DataFrame(columns =header_list).set_index("No")
    Instance_SourceMachine = pd.DataFrame(columns = header_list).set_index("No")
    Instance_CopiedFiles = pd.DataFrame(columns = header_list).set_index("No")
    Instance_Assign = pd.DataFrame(columns = header_list).set_index("No")
    Instance_NaibuFiles = pd.DataFrame(columns = header_list).set_index("Instance Name")
    Instance_JobFile = pd.DataFrame(columns = header_list).set_index("No")
    Instance_JobID = pd.DataFrame(columns = header_list).set_index("No")
    
    #Relation Tables
    relation_header = ["No", "ClassX instance", "ClassY instance", "Line Type", "Line Color", "Relation Type", "Label"]
    file_proid = pd.DataFrame(columns = relation_header).set_index("No")
    proid_aut = pd.DataFrame(columns = relation_header).set_index("No")
    proid_sm = pd.DataFrame(columns = relation_header).set_index("No")
    ffcopy = pd.DataFrame(columns = relation_header).set_index("No")
    ppcall = pd.DataFrame(columns = relation_header).set_index("No")
    file_rhis = pd.DataFrame(columns = relation_header).set_index("No")
    ffwrite = pd.DataFrame(columns = relation_header).set_index("No")
    fi_nai = pd.DataFrame(columns = relation_header).set_index("No")
    jf_jid = pd.DataFrame(columns = relation_header).set_index("No")
    jid_asn = pd.DataFrame(columns = relation_header).set_index("No")
    jid_xqt = pd.DataFrame(columns = relation_header).set_index("No")
    
    
    def excelize_cobollist(self,cobollist):     #Prepares necessary instance and relation tables to be written to excel 
        
        
        #Instances
        for cobolfile in cobollist:
            
            #File names to File
            row = pd.DataFrame([[len(self.Instance_ProgramFile) + 1,cobolfile.filename, "", "", "", ""]],
                               columns = self.header_list).set_index("No")
            self.Instance_ProgramFile = self.Instance_ProgramFile.append(row)
            #Files copied to CopiedFiles
            if len(cobolfile.copy_files)>0:
                for i in cobolfile.copy_files:
                    row = pd.DataFrame([[len(self.Instance_CopiedFiles) + 1,i, "", "", "", ""]],
                                       columns = self.header_list).set_index("No")
                    self.Instance_CopiedFiles = self.Instance_CopiedFiles.append(row)
            #Selected files to NaibuFiles
            if len(cobolfile.kanren_files)>0:
                for i in cobolfile.kanren_files:
                    row = pd.DataFrame([[len(self.Instance_NaibuFiles) + 1,i, "", "", "", ""]],
                                       columns = self.header_list).set_index("Instance Name")
                    self.Instance_NaibuFiles = self.Instance_NaibuFiles.append(row)
            self.Instance_NaibuFiles.drop_duplicates(inplace = True) #We do this here because we need to drop Assigned files before writing to file, and dropping duplicates makes that easier
            #Programs called to ProgramID
            if len(cobolfile.kanren_programs)>0:
                for i in cobolfile.kanren_programs:
                    row = pd.DataFrame([[len(self.Instance_ProgramID) + 1,i, "", "", "", ""]],
                                       columns = self.header_list).set_index("No")
                    self.Instance_ProgramID = self.Instance_ProgramID.append(row)
            #########################
            #Files written to NaibuFiles
            #if len(cobolfile.write_files)>0:
             #   for i in cobolfile.write_files:
              #      row = pd.DataFrame([[len(self.Instance_ProgramFile) + 1,i, "", "", "", ""]], columns = self.header_list).set_index("No")
               #     self.Instance_ProgramFile = self.Instance_ProgramFile.append(row)
            ############################
            #Revision Dates to RevisionHistory
            if len(cobolfile.revise_date)>0:
                for i in cobolfile.revise_date:
                    row = pd.DataFrame([[len(self.Instance_RevisionHistory) + 1,i, "", "", "", ""]],
                                       columns = self.header_list).set_index("No")
                    self.Instance_RevisionHistory = self.Instance_RevisionHistory.append(row)
            #Author to Author
            row = pd.DataFrame([[len(self.Instance_Author) + 1,cobolfile.author, "", "", "", ""]],
                               columns = self.header_list).set_index("No")
            self.Instance_Author = self.Instance_Author.append(row)
            #Program ID to ProgramID
            row = pd.DataFrame([[len(self.Instance_ProgramID) + 1,cobolfile.program_id, "", "", "", ""]],
                               columns = self.header_list).set_index("No")
            self.Instance_ProgramID = self.Instance_ProgramID.append(row)
            #SourceMachine to SourceMachine
            row = pd.DataFrame([[len(self.Instance_SourceMachine) + 1,cobolfile.sources, "", "", "", ""]],
                               columns = self.header_list).set_index("No")
            self.Instance_SourceMachine = self.Instance_SourceMachine.append(row)
        #Instances finished
        
        #Relations
            #File-ProgramID
            row = pd.DataFrame([[len(self.file_proid)+1, cobolfile.filename, cobolfile.program_id, "","", "", ""]],
                               columns = self.relation_header).set_index("No")
            self.file_proid = self.file_proid.append(row)
            #Author-Program
            row = pd.DataFrame([[len(self.proid_aut)+1,cobolfile.author, cobolfile.program_id, "","", "", ""]],
                               columns = self.relation_header).set_index("No")
            self.proid_aut = self.proid_aut.append(row)
            #SourceMachine-Program
            row = pd.DataFrame([[len(self.proid_sm)+1, cobolfile.sources, cobolfile.program_id, "","", "", ""]],
                               columns = self.relation_header).set_index("No")
            self.proid_sm = self.proid_sm.append(row)
            #Copy relation program id- copyfile
            for i in cobolfile.copy_files:
                row = pd.DataFrame([[len(self.ffcopy)+1, cobolfile.program_id,i, "","", "copy", "copy"]],
                                   columns = self.relation_header).set_index("No")
                self.ffcopy = self.ffcopy.append(row)
            #Call relation
            for i in cobolfile.kanren_programs:
                row = pd.DataFrame([[len(self.ppcall)+1, cobolfile.program_id,i, "","", "call", "call"]],
                                   columns = self.relation_header).set_index("No")
                self.ppcall = self.ppcall.append(row)
            #Program-RevisionHistory
            for i in cobolfile.revise_date:
                row = pd.DataFrame([[len(self.file_rhis)+1, i,cobolfile.program_id, "","", "", ""]],
                                   columns = self.relation_header).set_index("No")
                self.file_rhis = self.file_rhis.append(row)
            #Write relation
            for i in cobolfile.write_files:
                row = pd.DataFrame([[len(self.ffwrite)+1, cobolfile.program_id,i, "","", "write", "write"]],
                                   columns = self.relation_header).set_index("No")
                self.ffwrite = self.ffwrite.append(row)
            #NaibuFiles relation
            for i in cobolfile.kanren_files:
                row = pd.DataFrame([[len(self.fi_nai)+1, cobolfile.program_id,i, "","", "select", ""]],
                                   columns = self.relation_header).set_index("No")
                self.fi_nai = self.fi_nai.append(row)
            
            
         #Relations finished      
        self.programcheck = 1     #Means cobollist is excelized

    def excelize_joblist(self, joblist):     #Prepares necessary instance and relation tables to be written to excel
        
        for jobfile in joblist:
        
        #Instances
            #Copied files
            #Copy句は無視
            #if len(jobfile.copy_files)>0:
            #    for i in jobfile.copy_files:
            #        row = pd.DataFrame([[len(self.Instance_ProgramID) + 1,i, "", "", "", ""]],
            #                           columns = self.header_list).set_index("No")
            #        self.Instance_ProgramID = self.Instance_ProgramID.append(row)
            #Assign files
            if len(jobfile.assign)>0:
                for i in jobfile.assign:
                    row = pd.DataFrame([[len(self.Instance_Assign) + 1,i, "", "", "", ""]],
                                       columns = self.header_list).set_index("No")
                    self.Instance_Assign = self.Instance_Assign.append(row)
            #Job ID
            row = pd.DataFrame([[len(self.Instance_JobID) + 1,jobfile.job_id, "", "", "", ""]],
                                       columns = self.header_list).set_index("No")
            self.Instance_JobID = self.Instance_JobID.append(row)
            #Job File
            row = pd.DataFrame([[len(self.Instance_JobFile) + 1,jobfile.filename, "", "", "", ""]],
                                       columns = self.header_list).set_index("No")
            self.Instance_JobFile = self.Instance_JobFile.append(row)
            #XQT file
            for i in jobfile.xqt:
                row = pd.DataFrame([[len(self.Instance_ProgramFile) + 1, i , "", "", "", ""]],
                                   columns = self.header_list).set_index("No")
                self.Instance_ProgramFile = self.Instance_ProgramFile.append(row)      
        #Relations
            #JobFile-JobID
            row = pd.DataFrame([[len(self.jf_jid)+1, jobfile.filename, jobfile.job_id, "","", "", ""]],
                               columns = self.relation_header).set_index("No")
            self.jf_jid = self.jf_jid.append(row)
            #JobID-Assign
            for i in jobfile.assign:
                row = pd.DataFrame([[len(self.jid_asn)+1, jobfile.job_id, i, "","", "", ""]],
                                   columns = self.relation_header).set_index("No")
                self.jid_asn = self.jid_asn.append(row)
            #XQT
            for i in jobfile.xqt:
                row = pd.DataFrame([[len(self.jid_xqt)+1, jobfile.job_id, i, "","", "", ""]],
                                   columns = self.relation_header).set_index("No")
                self.jid_xqt = self.jid_xqt.append(row)  
                
        self.jobcheck = 1 #Means Joblist is excelized
    

    def spit_excel(self,address):   #Create the excel
        self.clean_naibu()
        with pd.ExcelWriter(address) as writer:
            #Main tables
            self.classtable.to_excel(writer, sheet_name="Class")
            self.relationtable.to_excel(writer, sheet_name="Relation")
            
            if self.jobcheck == 1:  #Check if job analysis is run
                #Jobcheck needs to come before programcheck because file names that are in both Assign and NaibuFiles tables should be in Assign class, which is achieved by writing it earlier to the file 
                self.Instance_Assign.drop_duplicates().to_excel(writer, sheet_name = "Instance_Assign")
                self.Instance_JobID.drop_duplicates().to_excel(writer, sheet_name = "Instance_JobID")
                self.Instance_JobFile.drop_duplicates().to_excel(writer, sheet_name = "Instance_JobFile")
                
                self.jf_jid.drop_duplicates().to_excel(writer, sheet_name = "relation9")
                self.jid_asn.drop_duplicates().to_excel(writer, sheet_name = "relation10")
                self.jid_xqt.drop_duplicates().to_excel(writer, sheet_name = "relation11")
                
            if self.programcheck == 1:   #Check if program analysis is run
                #Instance tables
                self.Instance_ProgramFile.drop_duplicates().to_excel(writer, sheet_name="Instance_ProgramFile")
                self.Instance_NaibuFiles.drop_duplicates().to_excel(writer, sheet_name = "Instance_NaibuFiles")
                self.Instance_RevisionHistory.drop_duplicates().to_excel(writer, sheet_name='Instance_RevisionHistory')
                self.Instance_Author.drop_duplicates().to_excel(writer, sheet_name= "Instance_Author")
                self.Instance_ProgramID.drop_duplicates().to_excel(writer, sheet_name="Instance_ProgramID")
                self.Instance_SourceMachine.drop_duplicates().to_excel(writer, sheet_name="Instance_SourceMachine")
                self.Instance_CopiedFiles.drop_duplicates().to_excel(writer, sheet_name = "Instance_CopiedFiles")
                
                #Relation tables
                self.file_proid.drop_duplicates().to_excel(writer, sheet_name="relation1")
                self.proid_aut.drop_duplicates().to_excel(writer, sheet_name="relation2")
                self.proid_sm.drop_duplicates().to_excel(writer, sheet_name="relation3")
                self.ffcopy.drop_duplicates().to_excel(writer, sheet_name="relation4")
                self.ppcall.drop_duplicates().to_excel(writer, sheet_name="relation5")
                self.file_rhis.drop_duplicates().to_excel(writer, sheet_name="relation6")
                self.ffwrite.drop_duplicates().to_excel(writer, sheet_name="relation7")
                self.fi_nai.drop_duplicates().to_excel(writer, sheet_name = "relation8")
                
        
    def clean_naibu(self):
        for i in self.Instance_Assign.values:
            try:
                self.Instance_NaibuFiles.drop(labels = i[0], inplace = True)
            except:
                pass
            
    


# In[6]:


class Cobol:        #Main functions to analyze cobol files
    cobollist = list()    #Container for all Cobolfile objects
    joblist = list()      #Container for all Jobfile objects
    naibulist = set()
    
    def __doc__(self):
        return "Cobol falan"
    
    def __init__(self):
        pass
    
    def hyouji(self, mode):
        if mode == "job":
            print("Job List")
            for i in joblist:
                i.prep()
                print("\n")
        else:
            print("Program List")
            for i in self.cobollist:
                i.prep()
                print("\n")

    def program_scanner(self, address):     #All files in the folder should be programs. Analyzes each file with program_analysis
        self.cobollist = list()
        for file in os.listdir(address):
            if file.endswith(".txt"):
                direc = address + "/" + file
                with open(direc, "r") as dire:
                    a = Cobolfile(file[:-4])   #Create a Cobolfile object for each program file
                    self.program_analysis(dire,a)
                    if a.program_id != "":     #An empty program id means the scanned file was not a program 
                        a.kanren_files = list(a.kanren_files) #So that it is subscriptable when replacing NaibuFiles 
                        a.write_files = list(a.write_files) #So that it is subscriptable when replacing NaibuFiles
                        self.cobollist.append(a)
        if len(self.cobollist) == 0:
            return
        else:
            return self.cobollist

    def job_scanner(self, address):       #All files in the folder should be jobs. Analyzes each file with job_analysis
        self.joblist = list()
        for file in os.listdir(address):
            if file.endswith(".txt"):
                direc = address + "/" + file
                with open(direc, "r") as dire:
                    a = Jobfile(file[:-4])      #Create a Jobfile object for each job file
                    self.job_analysis(dire, a)
                    if a.job_id != "":      #Empty job id means the scanned file was not a job file
                        self.joblist.append(a)
        if len(self.joblist) == 0:
            return
        else:
            self.replace_naibu()
            return self.joblist

    def job_analysis(self, file, jobdoc):     #Reads each line to assign them to attributes in a Jobfile object.
        for line in file:
            toke = [i.strip() for i in SPLIT(';|,|\*|\n|\s',line) if i != ""] #Normal split function only allows one delimiter, so regex is used
            if toke[0] == "@RUN":
                jobdoc.job_id = toke[1]
            #elif toke[0] == "@COPY":        #Can be ignored 
               # jobdoc.copy_files.add(toke[-1][toke[-1].rfind("."):toke[-1].find("/")].strip("."))   
            elif toke[0] == "@ASG":
                if toke[1] != "UP":
                    jobdoc.assign.add(toke[-1].strip("."))   
                else:
                    jobdoc.assign.add(toke[-2].strip("."))
            elif toke[0] == "@USE":
                jobdoc.use_pairs.add((toke[-2].strip("."), toke[-1].strip(".")))
            elif toke[0] == "@XQT":
                slic = toke[-1].find("/")
                if slic > 0:
                    jobdoc.xqt.add(toke[1][:slic].strip("."))
                else:
                    jobdoc.xqt.add(toke[1].strip("."))
        for x in jobdoc.use_pairs:
            jobdoc.recur(x[0], jobdoc.use_pairs, jobdoc.go_to_main, self.naibulist)   #When a file is finished, recur function finds the corresponding real file names for the inner names used
        

    def program_analysis(self,file, coboldoc):     #Reads each line to fill the attributes of the Cobolfile object
        for line in file:
            try:
                if int(line[:6]) > 9999:    #Checks if the first 6 characters look like a date
                    date = list(line[:6])
                    date.insert(-2, "/")
                    date.insert(-5, "/")     #Adds slashes so that Pandas does not assume them to be integers
                    date = "".join(date)
                    coboldoc.revise_date.add(date) 
                    
            except:
                pass
            if len(line) < 7 or "*" in line[6]:   #Skip if the line is a comment line
                continue
                
            toke = [i.strip().strip("./") for i in line.split()]
            try: 
                int(toke[0])
                toke.pop(0)
            except:
                pass
            
            if "SELECT" in toke:   #"If in"  is used because it is not always possible to know what will come before the command in Cobol
                pos = toke.index("SELECT")
                coboldoc.kanren_files.add(toke[pos + 1])
            elif "AUTHOR" in toke:
                pos = toke.index("AUTHOR")
                coboldoc.author = toke[pos + 1]
            elif "CALL" in toke:
                pos = toke.index("CALL")
                coboldoc.kanren_programs.add(toke[pos + 1].strip("''").strip('"'))
            elif "PROGRAM-ID" in toke:
                pos = toke.index("PROGRAM-ID")
                coboldoc.program_id = toke[pos+1]
                
            elif "FD" in toke:
                pos = toke.index("FD")
                recent_fd = toke[pos + 1]
            elif "SOURCE-COMPUTER" in toke:
                pos = toke.index("SOURCE-COMPUTER")
                coboldoc.sources = toke[pos + 1]
            elif "DATA" in toke:
                try :
                    pos = toke.index("RECORD")
                    if toke.index("DATA") == pos - 1:
                        if toke[pos + 1] == "IS":
                            coboldoc.recorddict[recent_fd] = toke[pos + 2]
                        else:
                            coboldoc.recorddict[recent_fd] = toke[pos + 1]
                except:
                    pass
            elif "WRITE" in toke:
                pos = toke.index("WRITE")
                tr = coboldoc.find_write_file(toke[pos+1])
                if tr is not None:
                    coboldoc.write_files.add(tr)
            elif "COPY" in toke:
                pos = toke.index("COPY")
                coboldoc.copy_files.add(toke[pos + 1].strip())
          
    def replace_naibu(self):       #Replace kanren_files of Cobolfile objects with their real names, which can be found in Jobfile object.
        #The Job file which executes a program also containes all the real file names for the inner names used in the Cobol program.
        for jobfile in self.joblist:
            for xfile in jobfile.xqt:
                for cobol in self.cobollist:
                    if cobol.filename == xfile:
                        for pair in jobfile.go_to_main:
                            temp = cobol.kanren_files
                            for i in range(len(temp)):
                                if pair[0] == temp[i]:
                                    temp[i] = pair[1]
                            temp = cobol.write_files
                            for i in range(len(temp)):
                                if pair[0] == temp[i]:
                                    temp[i] = pair[1]
                


# In[7]:


def input_path(message):
    while True:
        path = input(message)
        try:
            files = os.listdir(path)
            text_file = False
            if len(files) > 0:
                for file in files:
                    if file.endswith(".txt"):
                        text_file = True
                        break
                if text_file:
                    break
                else:
                    print("No .txt files found in the directory")
            else:
                print(f"No files found in {path}, please retry")
        except FileNotFoundError:
            print("Path not found, please try again")
            continue
    return path



def profol():
    global prodir, main
    prodir = filedialog.askdirectory(parent = main)
    if prodir == "":
        return
    Label(main, text = f"Path chosen for programs: {prodir}").pack()
    print(os.listdir(prodir))

def jobfol():
    global jobdir, main
    jobdir = filedialog.askdirectory(parent = main)
    if jobdir == "":
        return
    Label(main, text = f"Path chosen for jobs: {jobdir}").pack()
    print(os.listdir(jobdir))
    
def error():
    messagebox.showerror("An error has occured","Analysis failed. Please make sure you chose the correct folders for program and job files.\n(Current version supports unisys jobs only.)")

def run():
    global prodir, jobdir, main     
    if prodir == "" and jobdir == "":
        messagebox.showerror("Error with folders","You have not chosen either folder yet")
        return
    elif prodir == "":
        messagebox.showerror("Error with program folder","You have not chosen programs folder yet")
        return
    elif jobdir == "":
        messagebox.showerror("Error with jobs folder","You have not chosen jobs folder yet")
        return
    target_path = filedialog.askdirectory(parent = main)
    if target_path == "":
        return
    pro = Cobol()
    try:
        cobollist = pro.program_scanner(prodir)
        joblist = pro.job_scanner(jobdir)
    except: 
        error()
    if cobollist is None or joblist is None:
        error()
    Exob = Excelize()
    Exob.excelize_cobollist(cobollist)
    Exob.excelize_joblist(joblist)
    
    saved_to = (target_path + "\\Cobol analysis.xls")
    Exob.spit_excel(saved_to)
    messagebox.showinfo("Success",f"Analysis complete. File saved to: {saved_to}")
    
main = Tk()

prodir = ""
jobdir = ""
savedir = ""

Label(main, text = "First choose the directories:").pack()
Button(main, text = "Choose programs folder", command = profol).pack()
Button(main, text = "Choose jobs folder", command = jobfol).pack()
Label(main, text = "Then click the button below to run the analysis and save the result.").pack()
Button(main, text = "Run program", command = run ).pack()
    
main.mainloop()
    


            
