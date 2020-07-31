from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium import webdriver
import csv
from tkinter import *
import tkinter.messagebox
import requests,os,time
from zipfile import ZipFile

window = Tk()
window.title('Gitam Univeristy Semester Result')
window.geometry('500x520')
window.resizable(0,0)

live_status = StringVar()
up_frame = Frame(window, bg='black', height=20, width=500)
up_frame.pack(side=TOP, fill=X)

bottom_frame = Frame(window,height=20,width=500)
bottom_frame.pack(side=BOTTOM,fill=X)

frame2 = Frame(window, height=480, width=500,bg="gray")
frame2.pack(fill=X)

status = Label(bottom_frame,bd=1,relief=SUNKEN,anchor=W,textvariable=live_status)
status.pack(fill=X)

path = 'GITAM_logo.png'
if os.path.exists(path):
    pass
else:
    r=requests.get('https://www.gitam.edu/assets/images/GITAM_logo.png')
    with open('GITAM_logo.png','wb') as f:
        f.write(r.content)

photo=PhotoImage(file=path)
photo_label=Label(up_frame,image=photo,width=500,height=150)
photo_label.pack()



semester_label = Label(frame2, text="Semester   :", bg="white", fg="black",font="consolas 15")
section_label = Label(frame2, text="Section    :", bg="white", fg="black",font="consolas 15")


semester_entry= IntVar()
section_entry= IntVar()

semester_entry = Entry(frame2)
section_entry = Entry(frame2)

semester_label.place(x=100, y=95)
section_label.place(x=100, y=145)

semester_entry.place(x=250, y=100)
section_entry.place(x=250, y=150)

sec=0
gpa=[]
cgpa = []
roll=[]
name=[]
dmdw=[]
irs=[]
imp=[]
wt=[]
ai=[]
cd=[]
cd_lab=[]
wt_lab=[]
seminar=[]

dmdw_fail=0
irs_fail=0
imp_fail=0
wt_fail=0
ai_fail=0
cd_fail=0
cd_lab_fail=0
wt_lab_fail=0
seminar_fail=0

def search():
    maxStudents = 0
    passPercent = 0
    avgSGPA = 0
    avgCGPA = 0
    #Checking for empty fields
    if semester_entry.get() == '':
        tkinter.messagebox.showinfo('Error', "Semester value cannot be NULL.")
        return
    elif section_entry.get() == '':
        tkinter.messagebox.showinfo('Error', "Section value cannot be NULL.")
        return

    #else continue to conversion
    section_ip1 = int(section_entry.get())
    semester_ip1 =int(semester_entry.get())

    #check if any value exceeds it's limits
    if semester_ip1>8 :
        tkinter.messagebox.showinfo('Error',"Semester value cannot be greater than 8")
        semester_entry.delete(0, 'end')
        return
    elif section_ip1>19 :
        tkinter.messagebox.showinfo('Error',"Section value cannot be greater than 19")
        section_entry.delete(0,'end')
        return
    global sec
    sec = section_ip1
    sem = semester_ip1

    tkinter.messagebox.showinfo('NOTE','Running Chrome to get Results.Click --OK-- to proceed.')
    tkinter.messagebox.showinfo('NOTE','Please wait while we fetch data and record it. \nIMP: Donot close CHROME during operation.')
    # Initializing webdriver with its origin
    path = 'chromedriver.exe'
    if os.path.exists(path):
        pass
    else:
        live_status.set('Downloading chromedriver for automation...')
        window.update()
        time.sleep(5)
        r = requests.get('https://chromedriver.storage.googleapis.com/84.0.4147.30/chromedriver_win32.zip')
        with open('chromedriver.zip','wb') as f:
            f.write(r.content)
        live_status.set('Extracting chromedriver')
        window.update()
        time.sleep(5)
        zipfile = ZipFile('chromedriver.zip','r')
        zipfile.extractall()
        live_status.set('Extracted chromedriver, Opening chrome...')
        window.update()
    browser = webdriver.Chrome('chromedriver')

    # Automate to go to http address
    browser.get("https://doeresults.gitam.edu/onlineresults/pages/newgrdcrdinput1.aspx")

    for i in range(1,70):
        window.update()  # Roll loop
        if i == 69:
            browser.close()
            passPercent /=maxStudents
            passPercent *=100
            avgSGPA /= maxStudents
            avgCGPA /= maxStudents
            tkinter.messagebox.showinfo('Results','Students strength : '+str(maxStudents)+'\nPassPercentage: '+str(round(passPercent,2))+'\nAverege SGPA: '+str(round(avgSGPA,2))+'\nAverage CGPA: '+str(round(avgCGPA,2)))
            tkinter.messagebox.showinfo('NOTE',"The program has completed it's task, You can please press quit button to close the session")
            break
        semester = Select(browser.find_element_by_id("cbosem"))
        # Selecting semester option with Select class from Selenium
        semester.select_by_value(str(sem))
        # select_by_value selects semester with given value
        browser.find_element_by_id('txtreg').clear()
        searchBar = browser.find_element_by_id("txtreg")
        # find_element_by_id finds element and equalises it to other object
        if i < 10 and sec< 10:
            # send_keys sends string to a form in HTML
            searchBar.send_keys("12171030"+str(sec)+"00"+str(i))
        elif i < 10 and sec>= 10:
            searchBar.send_keys("1217103"+str(sec)+"00"+str(i))
        elif i >= 10 and sec < 10:
            searchBar.send_keys("12171030"+str(sec)+"0"+str(i))
        elif i >= 10 and sec>= 10:
            searchBar.send_keys("1217103"+str(sec)+"0"+str(i))
        elem = browser.find_element_by_id('Button1')
        elem.click()  # click() fun automates button clicks

        get_url = browser.current_url
        if get_url=='https://doeresults.gitam.edu/onlineresults/pages/newgrdcrdinput1.aspx':
            # browser.get("https://doeresults.gitam.edu/onlineresults/pages/newgrdcrdinput1.aspx")
            gpa.append("NO_SGPA_Recorded")
            cgpa.append("NO_CGPA_Recorded")
            roll.append('1217103130'+str(i))
            name.append("Page_Error")
            dmdw.append("Page_Error")
            irs.append("Page_Error")
            imp.append("Page_Error")
            wt.append("Page_Error")
            ai.append("Page_Error")
            cd.append("Page_Error")
            wt_lab.append("Page_Error")
            cd_lab.append("Page_Error")
            seminar.append("Page_Error")
            continue
        Name = browser.find_element_by_id("lblname")
        dmdw1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[2]/td[4]")
        irs1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[3]/td[4]")
        imp1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[4]/td[4]")
        wt1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[5]/td[4]")
        ai1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[6]/td[4]")
        cd1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[7]/td[4]")
        wt_lab1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[8]/td[4]")
        cd_lab1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[9]/td[4]")
        seminar1=browser.find_element_by_xpath("//*[@id='GridView1']/tbody/tr[10]/td[4]")
        marks=browser.find_element_by_id("lblgpa")
        marks1=browser.find_element_by_id("lblcgpa")
        regNo = browser.find_element_by_id("lblregdno")

        #TESTING FAILS AND SUMMING
        global dmdw_fail,irs_fail,imp_fail,wt_fail,ai_fail,cd_fail,cd_lab_fail,wt_lab_fail,seminar_fail
        if dmdw1.text=='F':
            dmdw_fail=dmdw_fail+1

        if irs1.text=='F':
            irs_fail=irs_fail+1

        if imp1.text=='F':
            imp_fail=imp_fail+1

        if wt1.text=='F':
            wt_fail=wt_fail+1

        if ai1.text=='F':
            ai_fail=ai_fail+1

        if cd1.text=='F':
            cd_fail=cd_fail+1

        if wt_lab1.text=='F':
            wt_lab_fail=wt_lab_fail+1
        if cd_lab1.text=='F':
            cd_lab_fail=cd_lab_fail+1
        if seminar1.text=='F':
            seminar_fail=seminar_fail+1

        # print(pd_fail,eem_fail,dbms_fail,flat_fail,se_fail,daa_fail,dbms_lab_fail,uml_lab_fail)
        # takes text of the HTML element and assigns it to variable
        dmdw.append(dmdw1.text)
        irs.append(irs1.text)
        imp.append(imp1.text)
        wt.append(wt1.text)
        ai.append(ai1.text)
        cd.append(cd1.text)
        wt_lab.append(wt_lab1.text)
        cd_lab.append(cd_lab1.text)
        seminar.append(seminar1.text)
        sgpa=marks.text
        gpa.append(sgpa)
        t_cgpa = marks1.text
        cgpa.append(t_cgpa)
        roll.append(regNo.text)
        name.append(Name.text)
        

        # back_elem = browser.find_element_by_id('Button1')
        # back_elem.click()
        browser.get("https://doeresults.gitam.edu/onlineresults/pages/newgrdcrdinput1.aspx")
        live_status.set('Extracted Roll no '+str(i))
        maxStudents = i
        if  float(sgpa)> 0:
            passPercent +=1
        avgSGPA += float(sgpa)
        avgCGPA += float(t_cgpa)

        # creates data_sec.csv and makes it read and write enable
        with open("data_"+str(sec)+".csv", "w+") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(["Name", "Roll No", "DMDW", "IRS","IMP","WT","AI","CD","WT LAB","CD LAB","SEMINAR","SGPA","CGPA"])
            for i in range(len(gpa)):
                # Writing fetched data into csv file
                writer.writerow([name[i], str(roll[i]),str(dmdw[i]),str(irs[i]),str(imp[i]),str(wt[i]),str(ai[i]),str(cd[i]),str(wt_lab[i]),str(cd_lab[i]),str(seminar[i]),str(gpa[i]),str(cgpa[i])])
            writer.writerow(["DMDW FAIL", "IRS FAIL","WT FAIL","AI FAIL","CD FAIL","WT FAIL","WT LAB FAIL","CD LAB FAIL","SEMINAR"])
            writer.writerow([dmdw_fail,irs_fail,imp_fail,wt_fail,ai_fail,cd_fail,wt_lab_fail,cd_lab_fail,seminar_fail])

#Submit_Button
submit_button = Button(frame2, text="Submit", fg="green", bg="white", command=search,height=1,width=8,font="consolas 15 bold")
submit_button.place(x=200,y=210)
def thank():
    tkinter.messagebox.showinfo('Thank Note','Thank You For Using the Product')
    exit()
quit_button=Button(text="Quit",fg="white",bg="red",command=thank,height=1,width=8,font="consolas 12 italic")
quit_button.place(x=400,y=450)

window.mainloop()
