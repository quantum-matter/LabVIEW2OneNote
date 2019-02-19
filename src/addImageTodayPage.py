'''
Author - Anurag Saha Roy
Date - October 2018

Add a new image to Today's page
'''

import requests
import urllib
import re

'''
Return a dict of various date items
library dependencies - datetime, calendar
'''

def variousDates():
    from datetime import date
    from datetime import datetime
    import calendar
    myDate=date.today()
    datesDict=dict()
    fullDate=datetime.now()
    dayName=calendar.day_name[myDate.weekday()]
    monthName=calendar.month_name[myDate.month]
    dayNum=myDate.day
    year=myDate.year
    datesDict.update({'fullDate':fullDate,'dayName':dayName, 'monthName':monthName, 'day':dayNum, 'year':year})
    return datesDict

'''
Return an access_token based on the supplied credentials
Please remember to update the four variables with respective credentials after the function definition below
'''
def getToken(tenantID, clientID, clientSecret):
    getTokenURL='https://login.microsoftonline.com/'+tenantID+'/oauth2/v2.0/token'
    getTokenHeaders={'Host':'login.microsoftonline.com', 'Content-Type':'application/x-www-form-urlencoded'}
    getTokenBody={'client_id':clientID, 'scope':'https://graph.microsoft.com/.default', 'client_secret':clientSecret,'grant_type':'client_credentials'}
    tokenResp=requests.get(getTokenURL, headers=getTokenHeaders, data=getTokenBody)
    access_token=tokenResp.json()['access_token']
    return access_token


_tenantID_=Enter Tenant ID
_clientID_=Enter Client ID
_clientSecret_=Enter Client Secret
email=Enter email ID
access_token=getToken(_tenantID_, _clientID_, _clientSecret_)


'''
function to get list of ALL pages
check if supplied page Title matches with any page in given Notebook
return PageID if found
return False as boolean if not found
function dependencies - NIL
'''
def getPageID(myTitle, myNotebook):

    getPagesURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/pages?expand=parentNotebook'

    getPagesHeaders={'Authorization': 'Bearer '+access_token, 'Accept':'application/json'}

    getPages=requests.get(getPagesURL, headers=getPagesHeaders)
    listOfPages=list(getPages.json()['value'])
    for page in listOfPages:
        if page['title']==myTitle and page['parentNotebook']['displayName']==myNotebook:
            hasPage = True
            pageID = page['id']
            return pageID
        else:
            hasPage = False
    return hasPage


'''
function to get HTML string of a page with supplied PageID
returns the content of the page as HTML string
returns error code 20112 for invalid ID
'''
def getPageData(myPageID):
    getPageDataURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/pages/'+myPageID+'/content?includeIDs=true'
    getPageDataHeaders={'Authorization': 'Bearer '+access_token, 'Accept':'application/json'}
    myPageData=requests.get(getPageDataURL, headers=getPageDataHeaders)
    return myPageData.text


'''
Fetch HTML content of the page supplied by myPageID
Find last occurence of div tag
Use a regular expression search to extract andn return last div ID
For new pages without div data, returns 'body'
'''
def findLastDiv(myPageID):
    pageHTML = getPageData(myPageID)
    try:
        lastDivIndex = pageHTML.rindex('div:')
        tempString =  pageHTML[lastDivIndex:]
        searchRes =  re.search(r'div:\S*}', tempString)
        return searchRes.group()
    except ValueError:
        return 'body'


'''
get NotebookID for the supplied section
return False as boolean if notebook doesn't exist
'''
def getNotebookID(myNotebook):
    getNotebooksURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/notebooks'
    getNotebooksHeaders={'Content-Type':'application/json', 'Authorization':'Bearer '+access_token}
    getNotebooks=requests.get(getNotebooksURL, headers=getNotebooksHeaders)
    listOfNotebooks = list(getNotebooks.json()['value'])
    #print(listOfNotebooks)
    for notebook in listOfNotebooks:
        if notebook['displayName'] == myNotebook:
            myNotebookID = notebook['id']
            hasNotebook= True
            break
        else:
            hasNotebook= False
    if hasNotebook:
        return myNotebookID
    else:
        return hasNotebook


'''get SectionGroupID for the supplied Notebook and Section
function dependencies - NIL
return the SectionGroupID if found
return False as boolean if the SectionGroup doesn't exist
'''
def getSectionGroupID(mySectionGroup, myNotebook):
    getSectionGroupsURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/sectionGroups'
    getSectionGroupsHeaders={'Content-Type':'application/json', 'Authorization':'Bearer '+access_token}
    getSectionGroups=requests.get(getSectionGroupsURL, headers=getSectionGroupsHeaders)
    listOfSectionGroups=list(getSectionGroups.json()['value'])
    #print(listOfSectionGroups)
    #hasSectionGroup=False
    for sectionGroup in listOfSectionGroups:
        if sectionGroup['parentNotebook'] == None:
                hasSectionGroup=False
                continue
        elif sectionGroup['parentNotebook']['displayName'] == myNotebook and sectionGroup['displayName'] == mySectionGroup:
                mySectionGroupID=sectionGroup['id']
                hasSectionGroup=True
                break
        else:
            hasSectionGroup=False
    if hasSectionGroup:
        return mySectionGroupID
    else:
        return hasSectionGroup


'''function to create a Section Group corresponding to a particular year in a given Notebook
function dependencies - getNotebookID() 
NOTE: must be run only after getSectionGroupID() has been run to check for the presence of sectionGroup
'''
def createSectionGroup(mySectionGroup, myNotebook):
    myNotebookID=getNotebookID(myNotebook)
    postSectionGroupURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/notebooks/'+myNotebookID+'/sectionGroups'
    postSectionGroupHeaders={'Content-Type':'application/json', 'Authorization':'Bearer '+access_token}
    postSectionGroupBody={'displayName': mySectionGroup}
    postSectionGroup=requests.post(postSectionGroupURL, headers=postSectionGroupHeaders, json=postSectionGroupBody)


'''
get SectionID for the supplied Notebook, sectionGroup and Section
create sectionGroup (year) if it doesn't exist
function dependencies - getNotebookID(), getSectionGroupID(), createSectionGroup()
return the SectionID if found
return False as boolean if the Section doesn't exist
'''
def getSectionID(mySection, mySectionGroup, myNotebook):
    mySectionGroupID=getSectionGroupID(mySectionGroup, myNotebook)
    if mySectionGroupID ==  False:
        createSectionGroup(mySectionGroup, myNotebook)
        mySectionGroupID=str(getSectionGroupID(mySectionGroup, myNotebook))
    getSectionsURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/sections'
    getSectionsHeaders={'Content-Type':'application/json', 'Authorization':'Bearer '+access_token}
    getSections=requests.get(getSectionsURL, headers=getSectionsHeaders)
    listOfSections= list(getSections.json()['value'])
    #print(listOfSections)
    myNotebookID=getNotebookID(myNotebook)
    for section in listOfSections:
        if section['parentNotebook']['id'] == myNotebookID and section['parentSectionGroup']['id'] == mySectionGroupID and section['displayName'] == mySection :
                mySectionID=section['id']
                hasSection=True
                break
        else:
            hasSection=False
    #print(mySectionID)
    #print(hasSection)
    if hasSection:
        return mySectionID
    else:
        return hasSection

'''
function to create a section corresponding to a particular month
function dependencies - getNotebookID(), getSectionGroupID()
NOTE: must be run only after getSectionID() has been run to check for the presence of section & sectionGroup
NOTE: Code breaks if run before getSectionID() when the specific SectionGroup doesn't exist
'''
def createSection(mySection, mySectionGroup, myNotebook):
    myNotebookID=getNotebookID(myNotebook)
    mySectionGroupID=getSectionGroupID(mySectionGroup, myNotebook)
    postSectionURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/sectionGroups/'+mySectionGroupID+'/sections'
    postSectionHeaders={'Content-Type':'application/json', 'Authorization':'Bearer '+access_token}
    postSectionBody={'displayName': mySection}
    postSection=requests.post(postSectionURL, headers=postSectionHeaders, json=postSectionBody)
    #print(postSection.text)


'''
function to create a blank OneNote page for the present Day
Takes as input the Notebook in which to create the new page for today
Assumes Notebook already exists
Pass Notebook Name as parameter while calling python program
Hierarchy: Year -> Month -> <DayOfWeek, Month Date, Year> as in 2018/July/Tuesday, July 10, 2018
Will create sectionGroup(year) if it doesn't exist
Will create section(month) if it doesn't exist
library dependencies - requests (external), urllib, datetime, calendar
function dependencies - variousDates(), getSectionID(), createSection(), getToken()
returns the pageID of today's page
'''
def createTodayPage(myNotebook):
    #access_token=getToken(_tenantID_, _clientID_, _clientSecret_)

    allDates=variousDates()
    pageName=str(allDates['dayName'])+', '+str(allDates['monthName'])+' '+str(allDates['day'])+', '+str(allDates['year'])
    
    if getPageID(pageName, myNotebook) == False:
        mySection=allDates['monthName']
        mySectionGroup=str(allDates['year'])
        mySectionID=(getSectionID(mySection, mySectionGroup, myNotebook))
        if mySectionID == False:
            createSection(mySection, mySectionGroup, myNotebook)
            mySectionID =  str(getSectionID(mySection, mySectionGroup, myNotebook))
        postPageURL = 'https://graph.microsoft.com/v1.0/users/'+email+'/onenote/sections/'+mySectionID+'/pages'
        postPageHeaders={'Content-Type':'application/xhtml+xml', 'Authorization':'Bearer '+access_token}
        postPageBody='''<!DOCTYPE html>
        <html>
          <head>
            <title>'''+pageName+'''</title>
            <meta name="created" content="'''+str(allDates['fullDate'])+'''"/>
          </head>
          <body>

          </body>
        </html>'''
        postPage =  requests.post(postPageURL, headers=postPageHeaders, data=postPageBody)
        return getPageID(pageName, myNotebook)
    else:
        return getPageID(pageName, myNotebook)


'''function to add image content to a given Page
takes PageID and image file name as input
image must be saved in the folder which is being hosted on localhost and tunneled 
Returns PATCH Response(requests object) as output
obtains ngrok web URL for accessing local image by calling ngrok REST API
'''
def updatePageImage(myPageID, imageName):
    ngrokResp=requests.get('http://localhost:4040/api/tunnels')
    ngrokHTTPS = ngrokResp.json()['tunnels'][1]
    ngrokURL=ngrokHTTPS['public_url']
    
    patchPageURL='https://graph.microsoft.com/v1.0/users/'+email+'/onenote/pages/'+myPageID+'/content?includeIDs=true'
    patchPageHeaders={'Content-Type':'application/json', 'Authorization':'Bearer '+access_token}
    patchPageBody=[
    {
        'target':findLastDiv(myPageID),
        'action':'append',
        'position':'after',
        'content':'</br><img src="'+ngrokURL+'/'+imageName+'" alt="New image from a localhost" />'
    }
    ]
    patchResponse = requests.patch(patchPageURL, headers=patchPageHeaders, json=patchPageBody)
    return patchResponse


'''
function to fetch pageID for today and add image
takes as input myNotebook
function dependencies - getToken(), variousDates(), getPageID(), updatePageImage()
'''
def addImageTodayPage(myNotebook):
    allDates=variousDates()
    pageName=str(allDates['dayName'])+', '+str(allDates['monthName'])+' '+str(allDates['day'])+', '+str(allDates['year'])
    myPageID=getPageID(pageName, myNotebook)

    if  myPageID == False:
        newPageID = createTodayPage(myNotebook)
        resp=updatePageImage(newPageID, 'tempImg.png')
    else:
        resp=updatePageImage(myPageID, 'tempImg.png')
        return resp

if __name__ == '__main__':
#MUST PROVIDE NOTEBOOK NAME AS ARGUMENTS TO main()
    import sys
    myNotebook = sys.argv[1]
    print("Please wait while image gets uploaded and the notebook is synced.")
    resp=addImageTodayPage(myNotebook)