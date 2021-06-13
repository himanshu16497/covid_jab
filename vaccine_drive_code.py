#bug1: the curser is becoming active whilst switching between tabs: solved by using altleft instead of alt
#[a for aadhar, v for voter card, p for passport, d for driving license]
#

import pyautogui as pyo
import time
import pandas as pd

data_path='C:\\Users\\Himanshu Bansal\\Desktop\\vaccine_details_45.xlsx'
index_path='C:\\Users\\Himanshu Bansal\\Desktop\\index_file_45.xlsx'

x1=727
y1=495

x2=599
y2=410


#index=pd.read_excel(data_path,engine='openpyxl')
#details=pd.read_excel(index_path,engine='openpyxl')

def switch_tab_once():
    pyo.keyDown('altleft')
    pyo.press('tab')
    pyo.keyUp('altleft')
    #pyo.press('enter')    
    return

def press_pageup():
    pyo.press('pgup')
    time.sleep(.5)
    pyo.press('pgup')
    time.sleep(.5)
    

def switch_tab_twice():
    pyo.keyDown('altleft')
    pyo.press('tab')
    pyo.press('tab')
    pyo.keyUp('altleft')
    #pyo.press('enter')    
    return

def copy():
    #pyo.hotkey('ctrl','a')
    pyo.hotkey('ctrl', 'c')
    time.sleep(.5)
    return

def paste():
    pyo.hotkey('ctrl', 'v')
    time.sleep(1)
    return

def move_right():
    pyo.press('right')
    time.sleep(.5)
    return

def move_left():
    pyo.press('left')
    time.sleep(.5)

def press_tab(n):
    for x in range(0,n):
        pyo.press('tab')
        time.sleep(.25)
        
def press_alpha(alpha):
    pyo.press(alpha)
    time.sleep(.25)
    
def write_string(string):
    string=str(string)
    pyo.write(string)
    time.sleep(.25)
  
   
def click_dose1():
    pyo.click(x2,y2)
    time.sleep(.5)

def click_add():
    pyo.click(x1,y1)
    time.sleep(1)



def update_index(index,index_path, index_value):
    print(type(index.iloc[0,0]))
    print(type(index_value))
    index.iloc[0,0]=index.iloc[0,0]+1
    print(index.iloc[0,0])
    print(index)
    #index.iloc[0,0]=index_value+1
    index.to_excel(index_path, engine='openpyxl',index=False)
    
def press_backspace():
    pyo.press('backspace')
    time.sleep(.5)

index=pd.read_excel(index_path,engine='openpyxl')
details=pd.read_excel(data_path,engine='openpyxl')

index_value=index.iloc[0,0]
#print(index_value)
#print(type(index_value))
user_details=details.iloc[index_value,:]

switch_tab_once()
time.sleep(1)
click_add()

click_dose1()

#press_pageup()  
#click_dose1()

press_tab(1)    #from the 1st dose radiobutton to document dropdown
press_alpha('a')

press_tab(1)    # at the addhar number input
write_string(user_details.AADHAR) # aadhar card number entered
print(type(user_details.AADHAR))

press_backspace()
write_string(str(user_details.AADHAR)[11])


press_tab(1)   #to choosing citizen dropdown
press_alpha('c')
time.sleep(1)    #let the dropdown appear

press_tab(1)
press_alpha('c') # choose citizen again

press_tab(1)
write_string(user_details.NAME)  # enter the user's name

  #enter the user's year of birth


if user_details.GENDER in ['F','f']:  #accordingly check the gender and choose
    press_tab(1)
    move_right()
elif user_details.GENDER in ['M','m']:
    press_tab(1)
    move_right()
    move_left()
    
press_tab(1)
write_string(user_details.DOB)

press_tab(1)
write_string(user_details.MOBILE)   #put the mobile number
press_tab(1)

update_index(index, index_path, index_value)



    




