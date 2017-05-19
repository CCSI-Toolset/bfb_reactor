###############################################################################
#           Automated Test of CCSI Bubbling Bed Simulator
#           Runs on Aspen Customer Modeler 7.3
#           Uses Python 2.7, pywinauto 0.4.3
#           Gregory Pope, LLNL
#           March 29, 2013
#           For Acceptance Testing Bubbling Fluidized Bed Reactor Model v0.10.c2
#           V1.1
#############################################################################
import pywinauto
import time


def ChangeMouse(script):
     #Assumes Ghost mouse is visable on desktop and scripts are in My Documents folder
     #Change(or load intital) mouse scripts into Ghost Mouse
     #Allows Ghost Mouse to change scripts during automated test run to create numerous mouse behaviors 
     print('Change Mouse Scripts')
     w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'GhostMouse 3.2', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     
     window.MenuItem(u'&File').Click()
     window.MenuItem(u'&File->&Open').Click()
     w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Open file...', class_name='#32770')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['ComboBox']
     ctrl.SetFocus()
     ctrl.ClickInput()
     ctrl.TypeKeys(script)
     ctrl.TypeKeys('{ENTER}')
     return

def SetMouseSpeed(speed):
     #Assumes Ghost mouse is visable on desktop
     #Change(or load intital) mouse speed setting into Ghost Mouse (speed is non-zero integer -10 to 10, 10 = 10 times faster, -10 1/10th speed)
     #Allows Ghost Mouse to change speeds during automated test run to create numerous mouse behaviors 
     w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'GhostMouse 3.2', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     window.MenuItem(u'&Options').Click()
     window.MenuItem(u'&Options->&Playback').Click()
     window.MenuItem(u'&Options->&Playback->Spee&d').Click()
     w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Speed setting', class_name='AutoIt v3 GUI')[0])
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['Trackbar']
     ctrl.SetFocus()
     if (speed == 10): ctrl.TypeKeys('{END}')
     elif (speed == -10): ctrl.TypeKeys('{HOME}')
     elif (speed == 1): ctrl.TypeKeys('{END}'+'{LEFT 9}')
     elif (speed == 2): ctrl.TypeKeys('{END}'+'{LEFT 8}')
     elif (speed == 3): ctrl.TypeKeys('{END}'+'{LEFT 7}')
     elif (speed == 4): ctrl.TypeKeys('{END}'+'{LEFT 6}')
     elif (speed == 5): ctrl.TypeKeys('{END}'+'{LEFT 5}')
     elif (speed == 6): ctrl.TypeKeys('{END}'+'{LEFT 4}')
     elif (speed == 7): ctrl.TypeKeys('{END}'+'{LEFT 3}')
     elif (speed == 8): ctrl.TypeKeys('{END}'+'{LEFT 2}')
     elif (speed == 9): ctrl.TypeKeys('{END}'+'{LEFT 1}')
     elif (speed == -9): ctrl.TypeKeys('{HOME}'+'{RIGHT 1}')
     elif (speed == -8): ctrl.TypeKeys('{HOME}'+'{RIGHT 2}')
     elif (speed == -7): ctrl.TypeKeys('{HOME}'+'{RIGHT 3}')
     elif (speed == -6): ctrl.TypeKeys('{HOME}'+'{RIGHT 4}')
     elif (speed == -5): ctrl.TypeKeys('{HOME}'+'{RIGHT 5}')
     elif (speed == -4): ctrl.TypeKeys('{HOME}'+'{RIGHT 6}')
     elif (speed == -3): ctrl.TypeKeys('{HOME}'+'{RIGHT 7}')
     elif (speed == -2): ctrl.TypeKeys('{HOME}'+'{RIGHT 8}')
     else: print ('speed must be non-zero integer -10 to 10')
     ctrl = window['Ok']
     ctrl.Click()
     return

#File needed to do simualtion in Aspen Custom Modeler/AM_Untitled
Needed_file='BFBv0.10.c2.acmf'
#assure we do not type too fast over remore set up
keywait =.5
#Added iterations to non-linear solver
solver_iters = 200

pwa_app = pywinauto.application.Application()

# Assure Ghost Mouse help application in tray using file drag_bubbling_bed.rms
raw_input('Assure Ghost Mouse app open on desktop, hit enter')


# Start Application
pwa_app.start_("C:\Program Files (x86)\AspenTech\AMSystem V7.3\Bin\AspenModeler.exe")
print ('Started Application')


#Dismiss Registration if present
print ('Dismiss Registration')
if (pywinauto.timings.WaitUntilPasses(10,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0])):
     w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Technology Product Registration', class_name='#32770')[0]
     window = pwa_app.window_(handle=w_handle)
     window.SetFocus()
     ctrl = window['Button2']
     ctrl.Click()

#Increase max iterations for non-linear solver
print ('Increase max iterations for non-linear solver')
w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Run->Solver Options
window.TypeKeys('%r')
time.sleep(keywait)
window.TypeKeys('v')
time.sleep(keywait)

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Solver Options', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['TabControl']
ctrl.SetFocus()
ctrl.Select(5)
time.sleep(keywait)
ctrl.TypeKeys('%i')
time.sleep(keywait)
window.TypeKeys(str(solver_iters))
time.sleep(keywait)
ctrl.TypeKeys('{ENTER}')
time.sleep(keywait)



#Select and Configure Physical Properties
print ('Select Physical Properties')
w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.Maximize()
window.TypeKeys('%t')
window.TypeKeys('g')

w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['RadioButton']
ctrl.Click()
ctrl = window['Button2']
ctrl.Click()

print ('Select Chemicals')

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Components Specifications - Data Browser]')[0])
window = pwa_app.window_(handle=w_handle)
window.edit.TypeKeys('CO2\r') # for carbon dioxide
window.edit.TypeKeys('H2O\r') # for water
window.edit.TypeKeys('N2\r')  # for Nitrogen

w_handle = pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Components Specifications - Data Browser]')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')

# Select Property Method
print ('Select Property Method')
ctrl = window['AfxOleControl9023']
ctrl.SetFocus()
ctrl = window['ComboBox14']
ctrl.Click()
ctrl.Select('PR-BM')
ctrl.Click()

#Next Step
#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')


w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Required Properties Input Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Required PROPS Input Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


w_handle = pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Economic Analysis', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Close']
ctrl.Click()

w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()


#Tools->Next
window.TypeKeys('%t')
window.TypeKeys('n')


w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Properties Plus Complete', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

print ('Exit Properties')
w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'PropsPlus.aprbkp - Aspen Properties V7.3 - aspenONE - [Control Panel]')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#File->Save
window.TypeKeys('%f')
window.TypeKeys('s')
time.sleep(.5)
# File->Exit
window.TypeKeys("%f")
window.TypeKeys("x")


#This code needed first time and window not disabled TODO 

#w_handle = pywinauto.findwindows.find_windows(title=u'Aspen Properties', class_name='#32770')[0]
#window = pwa_app.window_(handle=w_handle)
#window.SetFocus()
# Don't display anymore
#ctrl = window['&No']
#ctrl.Click()


w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Physical Properties Configuration', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()
print('Now move over chemicals')
#Select the Components List view in All Items
w_handle = pywinauto.timings.WaitUntilPasses(5,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
ctrl = window['TreeView']
ctrl.ClickInput()
time.sleep(2)
window.TypeKeys('{PGUP}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
#Works in both three across or two across width
window.TypeKeys('{RIGHT}')
window.TypeKeys('{RIGHT}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')

           
#Move the chemicals to the active window
w_handle = pywinauto.timings.WaitUntilPasses(15,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Build Component List - Default', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['List1']
ctrl.SetFocus()
ctrl.Select(0)
ctrl = window['>']
ctrl.Click()
ctrl = window['OK']
ctrl.Click()

print('Get the needed file')
#Import in the needed file
w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#File -> Import
window.TypeKeys('%f')
window.TypeKeys('i')


w_handle = pywinauto.timings.WaitUntilPasses(2,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Open', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['ComboBox']
ctrl.SetFocus()
ctrl = window['Edit']
ctrl.SetFocus()

#file name below may change, see note at top

window.edit.TypeKeys(Needed_file)
ctrl = window['&Open']
ctrl.Click()

print('Get the Adsorber model')
#Get the Adsorber model
w_handle = pywinauto.timings.WaitUntilPasses(6,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Exploring - Simulation']
ctrl.SetFocus()
time.sleep(.5)
#Narrow width for so upcoming ghost mouse script will work
ctrl.MoveWindow(x=0, y=0, width = 200)
ctrl = window['AfxMDIFrame902']
ctrl.SetFocus()
#Assure depth of window so mouse will work
ctrl.MoveWindow(x=0, y=0, height = 240)
ctrl = window['TreeView']
ctrl.SetFocus()
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{DOWN}')
window.TypeKeys('{TAB}')
window.TypeKeys('{RIGHT}')
window.TypeKeys('{TAB}')
window.TypeKeys('{TAB}')
window.TypeKeys('{DOWN}')
#Assure depth of window so mouse will work
ctrl = window['AfxMDIFrame903']
ctrl.SetFocus()
ctrl.MoveWindow(x=0, y=240, height = 605)

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

# Load correct mouse file
ChangeMouse('drag_bubbling_bed.rms')
SetMouseSpeed(1)

#Run Ghost Mouse Script to drag Model to Process Flowsheet (Control F2)
ctrl.TypeKeys('^{F2}')

#wait for mouse to do its thing
time.sleep(35)

#Rename icon
print ('Rename Icon')
#click on icon to assure menu bar correct
w_handle =pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
#ctrl.ClickInput()
ctrl.TypeKeys('^m')
w_handle =pywinauto.timings.WaitUntilPasses(3,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Input', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Edit']
ctrl.SetFocus()
ctrl.TypeKeys('BFBads'+'{ENTER}')
ctrl = window['OK']
ctrl.Click()


w_handle =pywinauto.timings.WaitUntilPasses(10,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')

#Set up the fixed device variables
print ('Set up the fixed device variables')
ctrl = window['BFBads.DeviceVariables Table']
ctrl.SetFocus()

ah = .8
Cr = 1.0
dPhx = .01
Dt = 9.0
dx = .02
emf = .5
fw = .2
hw = 1.5
K_d = 100
Lb = 5.0
lhx = .24
nor = 2500
SIType = 'Bottom'
SO_Type= 'Overflow'
Tref = 0
wthx = .003

window.TypeKeys(str(ah)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 3}')
time.sleep(keywait)
window.TypeKeys(str(Cr)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dPhx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dt)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(dx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(emf)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(fw)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(hw)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(K_d)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Lb)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(lhx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(nor)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 5}')
time.sleep(keywait)
window.TypeKeys('%{DOWN}'+ SIType +'{ENTER}'+'{LEFT}'+'{DOWN}')
time.sleep(2)
window.TypeKeys('%{DOWN}'+ SO_Type +'{ENTER}'+'{LEFT}'+'{DOWN 2}')
time.sleep(3)
window.TypeKeys(str(Tref)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 3}')
time.sleep(keywait)
window.TypeKeys(str(wthx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}')
time.sleep(keywait)
ctrl.Close()

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN 8}')
window.TypeKeys('{ENTER}')

#Set up the Sorbant variables
print ('Set up the Sorbant variables')
ctrl = window['BFBads.SorbentVariables Table']
ctrl.SetFocus()

A1 = .0562
A2 = 2.62
A3 = .1
dH1 = -52100
dH2 = -36300
dH3 = -64700
dS1 = -78.5
dS2 = -88.1
dS3 = -175
E1 = 28200
E2 = 58200
E3 = 57700
m1 = 1.17
nv = 2350
cps = 1.13
dp = .00015
F = 0
kp = 1.36
phis= 1.0
rhos = 442

window.TypeKeys(str(A1)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(A2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(A3)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dH1)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dH2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dH3)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dS1)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dS2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dS3)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(E1)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(E2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(E3)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(m1)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(nv)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(cps)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dp)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(F)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(kp)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(phis)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(rhos)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.Close()


# Setup Inlet Streams Tables
print('Setup Inlet Streams Tables')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN 4}')
window.TypeKeys('{ENTER}')

GasIN_F = 6000
GasIN_P = 1.3
GasIN_T = 43
GasIN_zCO2 =  0.13
GasIN_zH2O =  0.06
GasIN_zN2 =  0.81
HXIn_F = 60000
HXIn_P = 1.12
HXIn_T =33
HXIn_zCO2 =0
HXIn_zH2O =1
HXIn_zN2 =0
SolidIn_Fm = 600000
SolidIn_T = 65
SolidIn_wBic = .01
SolidIn_wCar = .7
SolidIn_wH2O = .7

window.TypeKeys(str(GasIN_F)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_P)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_T)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zCO2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zH2O)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zN2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_F)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_P)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_T)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_zCO2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_zH2O)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_zN2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_Fm)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_T)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_wBic)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_wCar)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_wH2O)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.Close()

# Move icon to center
print ('Move icon to center')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('d')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda: pywinauto.findwindows.find_windows(title=u'Find Object', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
#Find the icon
ctrl = window['ListBox']
ctrl.Click()
ctrl = window['Find']
ctrl.Click()
window.Close()
#click on icon to assure menu bar correct
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

#Run simulation, expect failure
print ('Run simulation, expect failure')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('{F5}')

print ('Dismiss error dialog box')
#Dismiss error dialog box
w_handle = pywinauto.timings.WaitUntilPasses(50,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Error', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


print ('Dismiss second error dialog box')
#Dismiss error dialog box
w_handle = pywinauto.timings.WaitUntilPasses(50,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Error', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


# Reset Simualtion
print ('Reset Simualtion')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys ('^{F7}')

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Reset', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['&Yes']
ctrl.Click()

# Click icon to assure correct menu
print ('Move icon to center')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

#Run IPSolve script
print ('Run IPSolve script')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
time.sleep(keywait)
window.TypeKeys ('p')
time.sleep(keywait)
window.TypeKeys('{ENTER}')

#Dismiss the done Dialog Box
print ('Dismiss the done dialog box')
w_handle = pywinauto.timings.WaitUntilPasses(320,1,lambda: pywinauto.findwindows.find_windows(title=u'Scripting', class_name='#32770')[0])                                          
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()


# Get temperature plot
print('Get temperature plot')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN 9}')
window.TypeKeys('{ENTER}')

print ('Test Results Pass Bubbling Bed Adsorber Aspen Modeler Simulation '+ Needed_file +' '+ time.asctime())



#######################################


#Regenerator

# Reset Simualtion
print ('Reset Simualtion')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys ('^{F7}')

w_handle = pywinauto.timings.WaitUntilPasses(20,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Reset', class_name='#32770')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['&Yes']
ctrl.Click()

#Rename icon
print ('Rename Icon')
#click on icon to assure menu bar correct
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()
ctrl.TypeKeys('^m')
w_handle = pywinauto.findwindows.find_windows(title=u'Input', class_name='#32770')[0]
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Edit']
ctrl.SetFocus()
ctrl.TypeKeys('BFBrgn'+'{ENTER}')
ctrl = window['OK']
ctrl.Click()

#Change Table 4.1.1 Fixed Device Variables

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN}')
window.TypeKeys('{ENTER}')

#Set up the fixed device variables
print ('Set up the fixed device variables')
ctrl = window['BFBrgn.InletStreams Table']
ctrl.SetFocus()


Dt = 7.0
dx = .03
Lb = 4.0
lhx = .12


window.TypeKeys(str(ah)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 3}')
time.sleep(keywait)
window.TypeKeys(str(Cr)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(dPhx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Dt)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(dx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(emf)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(fw)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(hw)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(K_d)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(Lb)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(lhx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(nor)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 5}')
time.sleep(keywait)
window.TypeKeys( '%{DOWN}'+ SIType +'{ENTER}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys('{DOWN 2}')
time.sleep(3)
window.TypeKeys(str(Tref)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 3}')
time.sleep(keywait)
window.TypeKeys(str(wthx)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}')              
ctrl.Close()


# Setup Inlet Streams Tables
print('Setup Inlet Streams Tables')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN 4}')
window.TypeKeys('{ENTER}')

GasIN_F = 10000
GasIN_P = 1.3
GasIN_T = 110
GasIN_zCO2 =  0.49
GasIN_zH2O =  0.49
GasIN_zN2 =  0.02
HXIn_F = 5000
HXIn_P = 6.9
HXIn_T =170
HXIn_zCO2 =0
HXIn_zH2O =1
HXIn_zN2 =0
SolidIn_Fm = 600000
SolidIn_T = 90
SolidIn_wBic = .5
SolidIn_wCar = 2.5
SolidIn_wH2O = 1.0

window.TypeKeys(str(GasIN_F)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_P)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_T)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zCO2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zH2O)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(GasIN_zN2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_F)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_P)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_T)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_zCO2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_zH2O)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(HXIn_zN2)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_Fm)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_T)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN 2}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_wBic)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_wCar)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
window.TypeKeys(str(SolidIn_wH2O)+ '{RIGHT}'+'%{DOWN}'+'Fixed'+'{LEFT}'+'{DOWN}')
time.sleep(keywait)
ctrl.Close()

# Click icon to assure correct menu
print ('Move icon to center')
w_handle = pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

#Invoke IPSolve script
#Run IPSolve script
print ('Run IPSolve script')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
time.sleep(keywait)
window.TypeKeys ('p')
time.sleep(keywait)
window.TypeKeys('{ENTER}')

#Dismiss the done Dialog Box
print ('Dismiss the done dialog box')
w_handle = pywinauto.timings.WaitUntilPasses(320,1,lambda: pywinauto.findwindows.find_windows(title=u'Scripting', class_name='#32770')[0])                                          
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['OK']
ctrl.Click()

# Get temperature plot
print('Get temperature plot')
w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
ctrl = window['Process Flowsheet Window']
ctrl.SetFocus()
ctrl.ClickInput()

w_handle =pywinauto.timings.WaitUntilPasses(1,0.5,lambda:pywinauto.findwindows.find_windows(title=u'Untitled - Aspen Custom Modeler V7.3 - aspenONE')[0])
window = pwa_app.window_(handle=w_handle)
window.SetFocus()
window.TypeKeys('%l')
window.TypeKeys('m')
window.TypeKeys('{DOWN 9}')
window.TypeKeys('{ENTER}')

print ('Test Results Pass Bubbling Bed Regenerator Aspen Modeler Simulation '+ Needed_file +' '+ time.asctime())
