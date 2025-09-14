*** Settings ***
Library    SeleniumLibrary
Library    ExcelLibrary
Resource    D:\\project_test\\projectTest\\registerkeyword\\RegiterKey.robot

*** Test Cases ***
Tc1: Register Users From Excel
    Open Application    ${url}    ${Browser}
    Go To link
    Open Excel Document    ${dataexcel}    register

    FOR    ${i}    IN RANGE    2    ${rows}+1
        Run Register Flow    ${i}
    END

    Save And Close Excel    ${dataexcel}
    Close Browser
----------------------------------------
*** Keywords ***
Open Application
    [Arguments]    ${url}    ${Browser}
    Open Browser    ${url}    ${Browser}
    Maximize Browser Window
    Set Selenium Speed    0.2
    Wait Until Element Is Visible    //p[@class='font-home1']
    Click Element    //a[text()='เข้าสู่ระบบ']
Go To link
    
    Click Element    //a[text()='สมัครสมาชิก']

Run Register Flow
    [Arguments]    ${i}
    
    ${run_flag}=    Read Excel Cell    ${i}    1    
    ${run_flag}=    Evaluate    '${run_flag}'.strip().upper()
    Run Keyword If    '${run_flag}' != 'Y'    Return From Keyword

    ${prefix}    Read Excel Cell    ${i}    3
    Run Keyword If    '${prefix}' != 'None'    Select From List By Value    //select[@id='pre_name']    ${prefix}

    ${Name}    Read Excel Cell    ${i}    4
    Input Text    //input[@name='fname']    ${Name}

    ${lastname}    Read Excel Cell    ${i}    5
    Run Keyword If    '${lastname}' == 'None' or '${lastname}' == ''    Set Variable    ${lastname}    ${EMPTY}
    Input Text    //input[@name='lastname']    ${lastname}

    ${major}    Read Excel Cell    ${i}    6
    Run Keyword If    '${major}' != 'None'    Select From List By Label    //select[@id='major']    ${major}

    ${Email}    Read Excel Cell    ${i}    7
    Input Text    //input[@id='email']    ${Email}

    ${number}    Read Excel Cell    ${i}    8
    Input Text    //input[@id='tell']    ${number}

    ${Username}    Read Excel Cell    ${i}    9
    Input Text    //input[@id='username']    ${Username}

    ${password}    Read Excel Cell    ${i}    10
    Input Password    //input[@id='password']    ${password}

    Click Button    //button[text()='สมัครสมาชิก']
    Sleep    2s

    Run Keyword And Ignore Error    Handle Success Alert

   ${Expected_Result}=    Read Excel Cell    ${i}      11
    ${flag}    ${actual_result}=    Check Result    ${Expected_Result}    ${i}

    ${status}=    Run Keyword If    ${flag}    Set Variable    pass    ELSE    Set Variable    fail

    #Write Excel Cell    ${i}    12    ${actual_result}

    Capture Page Screenshot    D:\\project_test\\projectTest\\screenshot_${status}_row_${i}.png
    
    IF    ${flag}
        Write Excel Cell   ${i}     13   pass
    ELSE
        Write Excel Cell   ${i}      13   fail
        Log To Console    Failed at row ${i}   
        Run Keyword And Ignore Error     Click Element    //button[contains(text(),'ยกเลิก')]
        Go To link

    END

Handle Success Alert
    ${alert_text}=    Handle Alert    action=ACCEPT
    Log To Console    ข้อความ popup: ${alert_text}
    [Return]    ${alert_text}


Check Result
    [Arguments]    ${Expected_Result}    ${i}
    ${popup_visible}=     Run Keyword And Return Status    Wait Until Element Is Visible    //button[@id='closePopup']    5s
    ${username_error}=    Run Keyword And Return Status    Wait Until Element Is Visible    //span[@id='usernameError']    5s
    ${password_error}=    Run Keyword And Return Status    Wait Until Element Is Visible    //span[@id='passwordError']    5s
    ${success_visible}=   Run Keyword And Return Status    Wait Until Element Is Visible    //p[contains(text(),'เว็บไซต์ขอทุนวิจัยของคณะวิทยาศาสตร์ มหาวิทยาลัยแม่')]    5s

    IF    ${popup_visible}
        ${actual_result}=    Get Text    //p[@class='font-main']
    ELSE IF    ${username_error}
        ${actual_result}=    Get Text    //span[@id='usernameError']
    ELSE IF    ${password_error}
        ${actual_result}=    Get Text    //span[@id='passwordError']
    ELSE IF    ${success_visible}
        ${actual_result}=    Get Text    //p[contains(text(),'เว็บไซต์ขอทุนวิจัยของคณะวิทยาศาสตร์ มหาวิทยาลัยแม่')]
    ELSE
        ${actual_result}=    Set Variable    ไม่สามารถตรวจสอบผลลัพธ์ได้
    END

    Log To Console    => Expected: ${Expected_Result}
    Log To Console    => Actual: ${actual_result}

    ${flag}=    Run Keyword And Return Status    Should Be Equal As Strings    ${Expected_Result}    ${actual_result}

    Run Keyword If    not ${flag}    Capture Page Screenshot    error_row_${i}.png

    [Return]    ${flag}    ${actual_result}

Save And Close Excel
    [Arguments]    ${dataexcel}
    Sleep    1s
    Save Excel Document    ${dataexcel}
    Close Current Excel Document
