# import gc
#
# threshold0 = gc.get_threshold()
# gc.set_threshold(0)
#import pywinauto
from pywinauto.application import Application
from pywinauto.application import Application
from pywinauto import Desktop, mouse,findwindows,keyboard
from pywinauto.controls.uia_controls import ComboBoxWrapper
from pywinauto.controls.win32_controls import EditWrapper
from pywinauto.controls.win32_controls import ListBoxWrapper
from ctypes import windll, Structure, c_long, byref
from pywinauto.keyboard import send_keys
import difflib
import random
import time, os
from pywinauto.keyboard import send_keys
import string
# gc.set_threshold(threshold0)

class MSAA:


    ROBOT_LIBRARY_SCOPE = "TEST SUITE"

    def __init__(self):
        '''On Inclusion of Library it automatically launches the Application LaunchPad'''
        Application().start("C:\\RBFG\\ru\\ru07\\RUAppBar.exe")
        appselect = None
        c=None
        app=None
        desktop=None
        alp=None
    def launchALP(self):
        ''' Method to Launch the Application Launch Pad'''
        time.sleep(3)
        a = Application(backend='uia').start("C:\\RBFG\\ru\\ru07\\RUAppBar.exe")
        self.alp=Desktop(backend='uia').window(best_match='Application Launch Pad')
    def launchApp(self,app):
        ''' Method to Launch the Application from within the Application Launch Pad'''
        time.sleep(3)
        self.a = Application().start("C:\\RBFG\\ru\\ru07\\RUAppBar.exe")
        self.alp.window(best_match=app,control_type="Button").click_input()

        if app=='Mortgages':
            app='Mortgage'
        handle='.* '+app+'.*'
        Desktop(backend="uia").window(title_re=handle).wait('visible')
        dlg= Desktop(backend="uia").window(title_re=handle)
        dlg.maximize()
        self.app = Desktop(backend="uia").window(title_re= handle)
    def clickLink(self,link):
        '''Method to Click Link. The name of the link is to match exactly to the input into this function.'''
        if '@' in link:
            field = link.split('@')
            search=field[1]+' '+field[0]
            a=self.app.descendants(control_type="Hyperlink")
            for i in a:
                if i.legacy_properties()['Name']==search:
                    i.click_input()


            print(field)
        else:
            self.app.child_window(title=link, control_type="Hyperlink").wait('visible')
            self.app.child_window(title=link, control_type="Hyperlink").click_input()
    def enterDate(self,**kwargs):
        '''Method to enter date fields'''
        print("I am in legacy")
        edit_controls = self.app.descendants(control_type="Edit")
        #print (edit_controls)
        for key,value in kwargs.iteritems():
           # print(key,value)
            if '@' in key:
                field=key.split('@')
                for i in self.app.descendants(control_type='HeaderItem'):
                    section=field[1]
                    fld=field[0]
                    self._entercomDate(section,fld,str(value),i)
            else:
                print("This is a test")
                self._gofilldate(edit_controls,key,value)
    def enterCMSDate(self,auto_id, **kwargs):
        '''Method to enter date fields'''
        print("I am in legacy")
        #edit_controls = self.app.child_window(auto_id=auto_id).descendants(control_type="Edit")
        #edit_controls = self.app.child_window(auto_id=auto_id, control_type="Edit")
        # self.app.child_window(auto_id="TempEndDate", control_type="Edit").set_text(date)
        # print (edit_controls)
        trytab = self.app.child_window(title_re, control_type="DataItem").parent()
        edit_controls = trytab.descendants(control_type="Edit")
        print(edit_controls)
        for key, value in kwargs.iteritems():
            print(key, value)
            # controliter = iter(edit_controls)
            newDate = value.split('/')
            for i in edit_controls:
                if key in i.text_block():
                    edit_controls.set_text(newDate[0])
                    edit_controls.Edit.type_keys("{TAB}")
                     #edit_controls.descendants.type_keys("{TAB}").set_text(newDate[1])
                    edit_controls.set_text(newDate[1])
                    edit_controls.next().set_text(newDate[2])
    def enterComboBox(self,option):
        a=self.app.ComboBox.rectangle()
        self.app.ComboBox.expand()
             #print(a)
        x=a.left+50
        if int(option)==0:
            print('1st run')
            y=a.bottom+10
        else:
            print('Running Now')
            y=a.bottom+22*(int(option)-1)
             #print(x,y)
             # time.sleep(2)
        mouse.click(button='left', coords=(x, y))
    def Select_From_ComboBox(self,title,option):
        #self.readscreen()
        #a=self.app.child_window(title="Goto:", control_type="ComboBox")
        #for itm in a:
        #    print(itm)
        self.app.child_window(control_type="ComboBox").select("Facility Details")
        #self.app.child_window(title=title, control_type="ComboBox").select(0)
        #self.app.child_window(title=title, control_type="ComboBox").send_keys("{DOWN 3}")
        #a = self.app.child_window(title=title, control_type="ComboBox").item_count()
        a = self.app.child_window(control_type="ComboBox").item_count()
        print(a)
        #self.app.child_window(title=title, control_type="ComboBox").
    def scroller(self,clicks):
        self.app.set_focus()
        self.app.draw_outline()
        rect = self.app.rectangle()
        coords = (random.randint(rect.left, rect.right), random.randint(rect.top, rect.bottom))
        # mouse.press(button='left', coords=coords)
        print("clicking at" + str(coords[0]) + str(coords[1]))
        mouse.scroll(coords=coords, wheel_dist=int(clicks))
    def _gofilldate(self,edit_controls,key,value):
        controliter=iter(edit_controls)
        print(list(controliter))
        #print("key is "+key+"value is"+value)
        date=value.split('/')
        if key == 'Goto' or key == 'Goto:':
                self.app[key].wait('visible')
            # #print("I am in kwargs")
                #print(key)
        for i in controliter:
            print(key, i.text_block)
            # if key == i.text_block(key(date[0])):
            #     controliter.next().set_text(date[1])
            #     controliter.next().set_text(date[2])
            #     return
            # if (difflib.SequenceMatcher(None, key, i.text_block()).ratio()) > 0.87:
            #     print(difflib.SequenceMatcher(None, key, i.text_block()).ratio())
            #     i.set_text(date[0])
            #     controliter.next().set_text(date[1])
            #     controliter.next().set_text(date[2])
            #     return
            if key in i.text_block():
                print('First')
                #print(i.text_block())
                #i.select(key)
                #self.app.child_window(title_re="Temporary End Day", control_type="Edit").set_text(date)
                i.set_text(date[0])
                controliter.next().set_text(date[1])
                controliter.next().set_text(date[2])
                return
    def clickGoto(self):
        time.sleep(1)
        mouse.click(button='left', coords=(70, 340))


    def Facilities(self):
        time.sleep(1)
        #mouse.click(button='left', coords=(60, 270))
        self.app.child_window(title="Facilities", control_type="Hyperlink", found_index=0).click_input()

    def Facilities_CI(self):
        time.sleep(1)
        # mouse.click(button='left', coords=(60, 270))
        self.app.child_window(title="Facilities", control_type="Hyperlink", found_index=1).click_input()

    def total_link(self, linknum):
        time.sleep(3)
        if int(linknum) == 0:
        #if linknum is 0:
            print('This is the first link')
            mouse.click(button='left', coords=(70, 388))
        elif int(linknum) == 1:
            self.app.child_window(title="Totals Not to Exceed (TNX)", control_type="Hyperlink", found_index=1).click_input()

    def Update_Borrower_CM(self):
        time.sleep(3)
        mouse.click(button='left', coords=(98, 270))

    def Select_A_Borrower(self):
        time.sleep(3)
        self.app.child_window(title="Select A Borrower", control_type="Hyperlink", found_index=0).click_input()

    def Select_A_Borrower_CI(self):
        time.sleep(3)
        self.app.child_window(title="Select A Borrower", control_type="Hyperlink", found_index=1).click_input()

    def Borrower_Fees_CM(self):
        time.sleep(3)
        self.app.child_window(title="Borrower Fees", control_type="Hyperlink", found_index=0).click_input()

    def Borrower_Fees_CI(self):
        time.sleep(3)
        self.app.child_window(title="Borrower Fees", control_type="Hyperlink", found_index=1).click_input()

    def total_link_CI(self):
        time.sleep(3)
        self.app.child_window(title="Totals Not to Exceed (TNX)", control_type="Hyperlink", found_index=1).click_input()

    def click_Facilities_selection_list_link(self):
        time.sleep(1)
        self.app.child_window(title_re="Facility Selection List", control_type="Hyperlink", found_index=0).click_input()

    def clickGoto2(self):
        time.sleep(1)
        mouse.click(button='left', coords=(87, 724))
        keyboard.SendKeys('OAV')
                    #625
    def _getcoord(self):

        class POINT(Structure):
            _fields_ = [("x", c_long), ("y", c_long)]

        def queryMousePosition():
            pt = POINT()
            windll.user32.GetCursorPos(byref(pt))
            windll.user32.GetCursorPos(byref(pt))
            return {"x": pt.x, "y": pt.y}

        pos = queryMousePosition()
        return  pos
    def enterFields(self,*args,**kwargs):
        '''Method to Enter texts into fields. Works both sequentially and specifically
        example: enterFields   123     abs     // This enters the fields sequentially from the first fieldi.e. 123 wil be entered in the first field and abs into the scond field so on and so forth
                enterFields     field1=123  field2=abs  //This enters 123 into field 1 and abs into field 2. field1 and field2 should be the names of the fields on the UI'''
        # self.app.#print_control_identifiers()

        for i,arg in enumerate(args):
            #print("I am in args")
            edit='Edit'+str(i+1)
            self.app[edit].set_text(arg)
        edit_controls = self.app.descendants(control_type="Edit")

        # edit_controls[0].set_text(110)

        for key,value in kwargs.iteritems():
            flag= True
            if key=='Goto' or key=='Goto:':
                self.app[key].wait('visible')
            if '@' in key:
                field = key.split('@')
                for i in self.app.descendants(control_type='HeaderItem'):
                    section = field[1]
                    fld = field[0]
                    self._entercomField(section, fld, str(value), i)
                continue
            for i in edit_controls:
                print(i.text_block())
                try:
                    if(difflib.SequenceMatcher(None, key.lower(), i.text_block().lower()).ratio())>0.87:
                         print('Key is'+key.lower())
                         print('test is'+i.text_block().lower())
                         print(difflib.SequenceMatcher(None, key.lower(), i.text_block().lower()).ratio())
                         i.set_text(value)
                         flag=False
                         break
                except:
                    if (i.text_block()=='^Loan Description:'):
                       print(i.text_block())
                       print(difflib.SequenceMatcher(None, 'loan description', i.text_block().lower()).ratio())
                       if (difflib.SequenceMatcher(None,'loan description', i.text_block().lower()).ratio()) > 0.87:
                          i.set_text(value)

            if (flag):
                for i in edit_controls:

                    #print(key.lower(), i.text_block().lower())
                    print('Test1')
                    if key.lower() in i.text_block().lower():
                        i.set_text(value)

    def _entercomField(self,section,field,value,i):
        if section.lower() == i.legacy_properties()['Name'].strip().lower():
             goup = i.parent()
             vals = goup.descendants(control_type='Edit')
             # try:
             self._searchEditField(section,field, value, vals)

    def _searchEditField(self,section,field,value,vals):
        # print(field)
        flag=False
        # date = date.split('/')
        # yr=None
        for i in self.app.descendants(control_type='HeaderItem'):
            if section.strip().lower()==i.legacy_properties()['Name'].strip().lower():
                goup=i.parent()
                vals = goup.descendants(control_type='Edit')
                controliter = iter(vals)
                for i in controliter:
                    val=i
                    if field in i.text_block():
                        i.set_text(value)
                        break


                        # controlite
                    # r.next().set_text(date[0])
                    # controliter.next().set_text(date[1])
                    # controliter.next().set_text(date[2])

                    # if(val.legacy_properties()['Name']=="^ISC Status / Date\r\n(YYYY / MM / DD): "):
                    #     val.draw_outline()
                    # yr = i
    def selectRadioButtons(self,*args):
        '''selects the radiobutton with the name. If multiple values are present for a radiobutton use field=<radiobutton>'''
        for i,arg in enumerate(args):
            if '=' in arg:
                s=arg.replace('=',' ')
                print(s)
            else:
                s=arg

            for i in self.app.descendants(control_type="RadioButton"):
                if s.lower() in i.legacy_properties()['Help'].lower():
                    print("gocha")

                    i.click_input()
    def enterTempDate(self,date):
        date = date.split('/')
        #self.app.child_window(title_re="Temporary End Day", control_type="Edit").set_text(date)
        self.app.child_window(auto_id="TempEndDate", control_type="Edit").set_text(date[0])
        self.app.child_window(title="Temporary End Month", control_type="Edit").set_text(date[1])
        self.app.child_window(title="Temporary End Day", control_type="Edit").set_text(date[2])

    def enterDates(self,autoId, month, day, date):
        date = date.split('/')
        #self.app.child_window(title_re="Temporary End Day", control_type="Edit").set_text(date)
        self.app.child_window(auto_id=autoId, control_type="Edit").set_text(date[0])
        self.app.child_window(title=month, control_type="Edit").set_text(date[1])
        self.app.child_window(title=day, control_type="Edit").set_text(date[2])

    def enterDates_Clear(self,autoId, month, day):
        self.app.child_window(auto_id=autoId, control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)
        self.app.child_window(title=month, control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)
        self.app.child_window(title=day, control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)

    def enterValues(self,autoId, title, value):
        value = value.split('/')
        self.app.child_window(auto_id=autoId, control_type="Edit").set_text(value[0])
        self.app.child_window(title=title, control_type="Edit").set_text(value[1])

    def clickButton(self,button):
        '''clicks the named button'''
        try:
            self.app.child_window(title_re=button, control_type="Image").click_input()
        except:
            self.app.child_window(title_re=button+' Button', control_type="Image").click_input()

    def Find_Def(self, linknum):
        time.sleep(3)
        if int(linknum) == 0:
            self.app.child_window(title="Find Button", control_type="Image", found_index=0).click_input()
        elif int(linknum) == 1:
            self.app.child_window(title="Find Button", control_type="Image", found_index=1).click_input()
        elif int(linknum) == 2:
            self.app.child_window(title="Find Button", control_type="Image", found_index=2).click_input()

    #def Find_Def(self,index):
        '''clicks on Find Button'''
        #self.app.child_window(title_re="Find Button", control_type="Image", found_index=index).click_input()

    def submit(self):
        '''clicks on Submit Button'''
        self.app.child_window(title_re="Submit Button", control_type="Image").click_input()
    def closePopUp(self):
        '''Closes the popup'''
        self.app.child_window(title="Cancel", control_type="Hyperlink").click_input()
    def acceptPopUp(self):
        '''Clicks on Yes within the popup'''
        self.app.child_window(title="Yes Button", control_type="Image").click_input()
    def rejectPopUp(self):
        '''Clicks on No within the Popup'''

        self.app.child_window(title="No Button", control_type="Image").click_input()
    def continuer(self):
        '''Clicks on Continue Button'''
        try:
            self.app.child_window(title="Continue Button", control_type="Image").click_input()
        except:
            mouse.click(button='left', coords=(140, 915))           #655
    def redisplay(self):
        '''Clicks on Redisplay Button'''
        self.app.child_window(title="Redisplay Button", control_type="Image").click_input()
    def back(self):
        self.app.child_window(title="Back Button", control_type="Image").click_input()
    def find(self):
        '''Clicks on Find Button'''
        self.app.child_window(title="Find button", control_type="Image").click_input()
    def Add(self):
        '''Clicks on Add Button'''
        self.app.child_window(title="Add button", control_type="Image").click_input()
    def go(self):
        '''Clicks on Go Button'''
        self.app.child_window(title="Go Button", control_type="Image").click_input()
        time.sleep(2)

    def Select_Available_Borrowing_Option(self, title):
        #set_item_focus(item)
        time.sleep(2)
        self.app.child_window(title=title, control_type="ListItem").click_input()

    def linkToSRF(self):
        '''Method to specificaly handle linktoSRF button in OLBB'''
        self.app.child_window(title="Link To SRF Button", control_type="Image").wait('visible')
        self.app.child_window(title="Link To SRF Button", control_type="Image").click_input()
    def findSRF(self):
        '''Clicks on Find SRF Button'''
        self.app.child_window(title="Find SRF Button", control_type="Image").wait('visible')
        self.app.child_window(title="Find SRF Button", control_type="Image").click_input()

    def Help(self):
        send_keys("{ESC}")

    def Print(self):
        send_keys("{ENTER}")

    def MERGE(self, value):
        send_keys(value)

    def Borrowing_option(self, value):
        self.app.child_window(title="Borrowing Options 2", control_type="Edit").click_input()
        time.sleep(1)
        send_keys(value)

    def Auth_Currency(self, value):
        self.app.child_window(title="Auth Currencies:", control_type="Edit").wait('visible')
        self.app.child_window(title="Auth Currencies:", control_type="Edit").click_input()
        time.sleep(2)
        send_keys("{TAB}")
        time.sleep(2)
        send_keys(value)


    def INT_ACC(self, value1, value2, value3, value4):
        send_keys(value1)
        time.sleep(2)
        send_keys("{TAB}")
        time.sleep(2)
        send_keys(value2)
        time.sleep(2)
        send_keys("{TAB}")
        time.sleep(2)
        send_keys(value3)
        time.sleep(2)
        send_keys("{TAB}")
        time.sleep(2)
        send_keys(value4)

    def User_Details(self, value1, value2, value3):
            send_keys(value1)
            time.sleep(2)
            send_keys("{TAB}")
            time.sleep(2)
            send_keys(value2)
            time.sleep(2)
            send_keys("{TAB}")
            time.sleep(2)
            send_keys(value3)


    def extractTextFromResults(self, *args):
        '''Extracts the text from Results frame '''
        edit_controls=None
        try:
            edit_controls = self.app.child_window(auto_id="results").descendants(control_type="Edit")
        except:
            trytab= self.app.child_window(title="Results", control_type="DataItem").parent()
            edit_controls = trytab.descendants(control_type="Edit")
        tables=self.app.descendants(control_type="Table")
        for i in edit_controls:
            print(i.text_block())
        for arg in args:
            if arg == 'Goto':
                self.app[arg].wait('visible')
            print(arg)
            for i in edit_controls:
                print(arg, i.text_block())
                if (difflib.SequenceMatcher(None, arg, i.text_block()).ratio()) > 0.87:
                    print(difflib.SequenceMatcher(None, arg, i.text_block()).ratio())
                    print('Returning :'+i.legacy_properties()['Value'])
                    return (i.legacy_properties()['Value']).strip()

            try:
                print("I AM HERERERE")
                tables=self._searchForTable(self.app.child_window(auto_id="results"))
                # print(headerList)
                # self._RetrieveTableItems(self.app.child_window(auto_id="results"))

            except:
                tables=self._searchForTable(self.app.child_window(title="Results", control_type="DataItem"))

            print(tables)
        return(tables)
            # print(tables[1][Field Number])
    def _searchForTable(self,reshandle):
        tables={}
        rowlist={}
        # rows={}
        rowcount=0
        headerList=[]
        dataList=[]
        head=None
        x = reshandle.descendants(control_type='Table')
        print(x)
        for pointer in x:
            if len(pointer.descendants(control_type="HeaderItem")) > 0:
                headers = pointer.descendants(control_type="HeaderItem")
                Items=pointer.descendants(control_type="DataItem")
                for i in headers:
                    head=i
                    headerList.append(i.legacy_properties()['Name'])
                for i in Items:
                    head=i
                    dataList.append(i.legacy_properties()['Name'])
                datait=iter(dataList)
                headerit=iter(headerList)
                while True:
                    tmptables={}
                    rows={}
                    try:
                        for x in range(len(headerList)):
                            tmp={}
                            rows[headerList[x]] = next(datait)
                            tmp=rows
                        rowlist[rowcount]=tmp
                        rowcount+=1

                        # rowlist.append(tmp)

                    except StopIteration:
                        return rowlist
                        break  # Iterator exhausted: stop the loop



    def _RetrieveTableItems(self,reshandle):
        print("I am trying to retrieve")
        tables = {}
        headerList = {}
        dataList = {}
        x = reshandle.descendants(control_type='Table')
        print(x)
        for pointer in x:
            if len(pointer.descendants(control_type="HeaderItem")) > 0:
                head=pointer.parent().descendants(control_type='DataItem')
                for i in head:
                    print(head.text_block())
            break
    def extractTextFromField(self,*args):
        edit_controls = self.app.descendants(control_type="Edit")
        for arg in args:
            if arg=='Goto':
                self.app[arg].wait('visible')
            #print("I am in args")
            #print(arg)
            for i in edit_controls:
                #print(arg, i.text_block())
                if(difflib.SequenceMatcher(None, arg, i.text_block()).ratio())>0.87:
                    return  i.legacy_properties()['Value']
    def enterCombos(self,**kwargs):
        combo_controls = self.app.descendants(control_type="ComboBox")
        text_controls=self.app.descendants(control_type="Text")
        for key, value in kwargs.iteritems():
             for x in text_controls:
                         if key in x.legacy_properties()['Name']:
                             # print(x.legacy_properties())
                             # print(x.descendants())
                             # print(x.rectangle())
                             x = x.parent()
                             # print(x.legacy_properties())
                             # print(x.descendants())
                             # print(x.rectangle())
                             x = x.parent()
                             # print(x.legacy_properties())
                             # for i in x.descendants():
                             #     print(i.legacy_properties())
                             x = x.parent()
                             # for i in x.descendants():
                             #     print(i.legacy_properties())
                             # x=x.parent()
                             for count, i in enumerate(x.descendants()):

                                 try:  # i.expand()
                                     a = i.rectangle()
                                     i.expand()
                                     print(value)
                                     x = a.left + 50
                                     if option == 1:
                                         y = a.bottom + 10
                                     else:
                                         y = a.bottom + 22 * (value - 1)

                                     print(x, y)
                                     mouse.click(button='left', coords=(x, y))


                                 except():
                                     continue
    def exitoption(self):
        self.app.child_window(title="Main Menu", control_type="Hyperlink").wait('visible')
        self.app.child_window(title="Main Menu", control_type="Hyperlink").click_in
        # self.app.#print_control_identifiers()put()
    def closeApp(self):
        #
        # self.app.window().#print_control_identifchoteiers()
        self.app.child_window(title="Close", control_type="Button").click_input()
    def readscreen(self):
        self.app.print_control_identifiers()
    def _entercomDate(self,section,field,date,i):
        if section.lower() == i.legacy_properties()['Name'].strip().lower():
             goup = i.parent()
             vals = goup.descendants(control_type='Edit')
             self.searchDateField(section,field, date, vals)
    def searchDateField(self,section,field,date,vals):
        # print(field)
        flag=False
        date = date.split('/')
        yr=None
        for i in self.app.descendants(control_type='HeaderItem'):
            if section.strip().lower()==i.legacy_properties()['Name'].strip().lower():
                goup=i.parent()
                vals = goup.descendants(control_type='Edit')
                controliter = iter(vals)
                for i in controliter:
                    val=i
                    print(i.text_block())
                    if(field in i.text_block()):
                        flag=True

                    # print("Next is  "+ controliter.next().legacy_properties()['Name'])
                    # if( field in val.legacy_properties()['Name']):
                    #     val.draw_outline()
                    #     yr=date[0]
                    #     tmp=controliter
                    if 'Month'.strip().lower() in i.legacy_properties()['Name'].strip().lower() and flag:
                            # print(controliter.next().legacy_properties()['Name]'])
                            yr.set_text(date[0])
                            i.set_text(date[1])
                            controliter.next().set_text(date[2])
                            break
                    yr = i
    def openActivateLoan(self,srf):
        self.app.child_window(auto_id="srf_numbers_radio", control_type="RadioButton").click_input()
        time.sleep(0.5)
        keyboard.send_keys(srf)
        self.app.child_window(title="Find SRF Button", control_type="Image").wait('ready')
        self.app.child_window(title="Find SRF Button", control_type="Image").click_input()

    def Borr_Fees(self, value1, value2, value3, value4):
        self.app.child_window(title="^Fee Type 1", control_type="Edit").click_input()
        send_keys(value1)
        time.sleep(1)
        self.app.child_window(title="^Frequency 1", control_type="Edit").click_input()
        send_keys(value2)
        time.sleep(1)
        self.app.child_window(title="^Currency 1", control_type="Edit").click_input()
        send_keys(value3)
        time.sleep(1)
        self.app.child_window(title="Fee Amount 1", control_type="Edit").click_input()
        send_keys(value4)

    def Borr_Fees_Clear(self):
        self.app.child_window(title="^Fee Type 1", control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)
        self.app.child_window(title="^Frequency 1", control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)
        self.app.child_window(title="^Currency 1", control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)
        self.app.child_window(title="Fee Amount 1", control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')

    def Select_Item(self):
        time.sleep(2)
        #self.app.child_window(control_type="RadioButton", found_index=5).click_input()
        next = self.app.child_window(title_re="Next", control_type="Hyperlink", found_index=0)
        #if next.is_enabled() == True:
        try:
            while next.is_visible() == True:
                print('Clicking Next')
                next.click_input()
                if next.is_visible() == False:
                    break
            self.app.child_window(control_type="RadioButton", found_index=0).click_input()
        except:
            #a = self.app.child_window(control_type="RadioButton").find_elements()
            a = self.app.descendants(control_type="RadioButton")
            self.app.child_window(control_type="RadioButton", found_index=(len(a) - 1)).click_input()

    def Select_Item_1(self):
        time.sleep(2)
        # self.app.child_window(control_type="RadioButton", found_index=5).click_input()
        checkbox = self.app.child_window(title_re="Include", control_type="CheckBox", found_index=0)
        # if next.is_enabled() == True:
        try:
            while checkbox.is_enabled() == True:
                print('Clicking Next')
                checkbox.click_input()
                if checkbox.is_enabled() == False:
                    break
            self.app.child_window(control_type="CheckBox", found_index=0).click_input()
        except:
            # a = self.app.child_window(control_type="RadioButton").find_elements()
            a = self.app.descendants(title_re="Exclude", control_type="CheckBox")
            self.app.child_window(title_re="Exclude", control_type="CheckBox", found_index=(len(a) - 1)).click_input()

    def Select_Item_Include(self):
            time.sleep(2)
            self.app.child_window(control_type="CheckBox", title="Include - 001-E").click_input()
            time.sleep(2)
            self.app.child_window(control_type="CheckBox", title="Include - 002-E").click_input()

    def Select_Item_Exclude(self):
            time.sleep(2)
            self.app.child_window(control_type="CheckBox", title="Exclude - *001-E").click_input()
            time.sleep(2)
            self.app.child_window(control_type="DataItem", title="SALKDHAKS").click_input()

    def Select_Item_Include_Exclude(self):
            time.sleep(2)
            self.app.child_window(control_type="CheckBox", title="Include - 004-C").click_input()
            time.sleep(2)
            self.app.child_window(control_type="CheckBox", title="Exclude - *002-E").click_input()

    def Click_Enter_On_DialogBox(self):
        '''clicks on Submit Button'''
        send_keys("{ENTER}")

    def Click_Cancel_On_DialogBox(self):
        '''clicks on Cancel Button'''
        send_keys("{TAB}")
        time.sleep(1)
        send_keys("{TAB}")
        time.sleep(1)
        send_keys("{ENTER}")

    def Clear_Main(self):
        time.sleep(2)
        self.app.child_window(title="SRF #:", control_type="Edit").click_input()
        send_keys('^a{BACKSPACE}')
        time.sleep(1)

    def Curr_Code(self, value):
        send_keys("{TAB}")
        time.sleep(2)
        send_keys(value)

    def Select_Item_To_Detach(self, acc):
        self.app.child_window(title=acc, control_type="RadioButton").click_input()

    def Attach_Account(self,*args,**kwargs):
        '''Method to Enter texts into fields. Works both sequentially and specifically
        example: enterFields   123     abs     // This enters the fields sequentially from the first fieldi.e. 123 wil be entered in the first field and abs into the scond field so on and so forth
                enterFields     field1=123  field2=abs  //This enters 123 into field 1 and abs into field 2. field1 and field2 should be the names of the fields on the UI'''
        # self.app.#print_control_identifiers()

        for i,arg in enumerate(args):
            #print("I am in args")
            edit='Edit'+str(i+1)
            self.app[edit].set_text(arg)
        edit_controls = self.app.descendants(control_type="Edit")

        # edit_controls[0].set_text(110)

        for key,value in kwargs.iteritems():
            flag= True
            if key=='Goto' or key=='Goto:':
                self.app[key].wait('visible')
            if '@' in key:
                field = key.split('@')
                for i in self.app.descendants(control_type='HeaderItem'):
                    section = field[1]
                    fld = field[0]
                    self._entercomField(section, fld, str(value), i)
                continue
            controliter = iter(edit_controls)
            account = value.split('-')
            print(account)

           # if key in i.text_block():
            for i in controliter:

                    #print(key.lower(), i.text_block().lower())
                print('Test1')
                if key in i.text_block():
                    i.set_text(account[0])
                    controliter.next().set_text(account[1])
                    #controliter.next().type_keys()
                    return

# o = MSAA()
# o.launchALP()
# o.launchApp('Credit Monitoring')
# a={}
# a['Approval Date(YYYY//MM//DD)'] = '2019/02/29'
# o.enterDate()
# o.app = Desktop(backend="uia").window(title_re='OLBB')
# o.alp.print_control_identifiers()


# o.print_control_identifiers()
# a={}
# a['^Client Number']='1'
# o.enterFields(a)
# o.clickButton('Validate')
# a={}
# a['Goto']='OVL'



# mouse.click(button='left', coords=(140, 655))
# o.clickLink('Proposal Inquiry')
# o.selectRadioButtons('Non-Personal')
# o.clickButton('Ok')
