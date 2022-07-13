*** Settings ***
Documentation     Template robot main suite.
Library           Collections
Library           MyLibrary
Variables         MyVariables.py
Library           RPA.Windows
Library    RPA.Desktop
Library           RPA.Browser.Selenium
# Library    RPA.Desktop
# Library    RPA.Desktop.Windows
# Variables
*** Tasks ***
Task1
    ${d1}    ${d2}    Get Schedule Week  
      
    Log    ${d1}
    Log    ${d2}
    # Log    ${file_l1}
    # Open Chrome Browser    https://smartoffice.mobifone.vn/eoffice/#/app/9627b6fa-be51-4c38-9580-8a9e355773e9/dashboard
    Set Download Directory    C:\\Users\\admin\\Desktop\\lichtuan    True
    Open Available Browser    https://smartoffice.mobifone.vn/eoffice/#/app/9627b6fa-be51-4c38-9580-8a9e355773e9/dashboard 
    
    Input Text    //*[@id="username"]    giamsat_spamvoice@mobifone.vn
    Input Password    //*[@id="password"]    mobifone123456@@
    Click Button    //*[@id="bth-login"]
    Sleep    5
    Input Text    //*[@id="main-otp"]/div/div/div/div/div/form/div[2]/div/div/div/div/input    854564
    Sleep    5   
    # Click Link    //*[@id="mainMenu"]/ul/a[2]
    # Sleep    5
     
    ${link}    Create Link Download    ${d1}    ${d2}
    ${current_time}    ${current_date}   Get Current Time
    Go To    ${link}
    # Click Button    //*[@id="subHeader"]/div[1]/button[2]
    # Sleep    3
    # Input Text    /html/body/div[2]/div/div[2]/div/section/div[1]/form/div/div/div/div[2]/div/div/div[2]    Trung tâm Công nghệ thông tin MobiFone
    # Click Element    /html/body/div[2]/div/div[2]/div/section/div[1]/form/div/div/div/div[2]/div/div/div[2]
    # https://smartoffice.mobifone.vn/BusinessService/report/lichTuanChuaDuyetReport?fromTime=1657472400&endTime=1657731600000&userDeptRoleId=9627b6fa-be51-4c38-9580-8a9e355773e9&exportType=processed&deptId=TTCNTT-TCT
    Sleep   5
    Log    ${current_time}
    Log    ${current_date}
    ${file_l1}    Get File Name    ${current_date}    ${current_time}
    ${path}    File Location   ${file_l1}    

    ${mlist}    Read File    ${path}
    FOR    ${element}    IN    @{mlist}
        Log    ${element}
    END

    RPA.Desktop.Press Keys    cmd  r
    RPA.Desktop.Type Text    outlookcal:
    RPA.Desktop.Press Keys    enter
    Sleep    2
    FOR    ${element}    IN    @{mlist}
        Event to Calendar    ${element}
        Sleep    0.5
    END
    
    RPA.Desktop.Press Keys    alt  f4
    # Add An Event to Calendar    ${mlist}[0]   


*** Keywords ***

Event to Calendar
    
    [Arguments]    ${element}
    RPA.Desktop.Press Keys    ctrl  n
    RPA.Windows.Click    name:"Event description"
    RPA.Desktop.Type Text    ${element}[4]
    Sleep    1.5
    RPA.Windows.Click    name:"Start Date"
    RPA.Desktop.Type Text    ${element}[0]
    Sleep    1.5
    RPA.Windows.Click    name:"Start time"
    RPA.Desktop.Type Text    ${element}[1]
    Sleep    1.5
    RPA.Windows.Click    name:"Location"
    RPA.Desktop.Type Text    ${element}[3]
    Sleep    1.5
    RPA.Windows.Click    name:"Event name"
    RPA.Desktop.Type Text    ${element}[2]
    Sleep    1.5
    RPA.Desktop.Press Keys    ctrl  s


    