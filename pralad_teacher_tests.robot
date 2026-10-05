*** Settings ***
Documentation    Individual browser tests for Pralad Raya S400417. Run only against a local disposable database.
Library          SeleniumLibrary
Suite Setup      Open Test Browser
Suite Teardown   Close Browser

*** Variables ***
${BASE_URL}          http://localhost:8000
${BROWSER}           chrome
${TEACHER_ID}        070101
${TEACHER_PASSWORD}  1234
${TEACHER_PIN}       185748

*** Test Cases ***
TC01 Teacher Login With Valid Credentials
    Log In As Teacher
    Location Should Contain    change-password.php

TC02 Invalid Teacher Login Is Rejected
    Go To    ${BASE_URL}/index.php
    Click Element    id:menu-student-signin
    Input Text    name:regno    ${TEACHER_ID}
    Input Password    name:password    incorrect-password
    Click Button    css:button[type="submit"]
    Wait Until Location Contains    index.php
    Click Element    id:menu-student-signin
    Page Should Contain    Invalid Teacher ID or Password

TC03 Protected Enrollment Redirects Without Login
    Go To    ${BASE_URL}/logout.php
    Go To    ${BASE_URL}/enroll.php
    Wait Until Location Contains    index.php

TC04 Incorrect PIN Is Rejected
    Log In As Teacher
    Go To    ${BASE_URL}/pincode-verification.php
    Input Password    name:pincode    000000
    Click Button    css:button[type="submit"]
    Page Should Contain    Wrong Pincode

TC05 Correct PIN Opens Enrollment
    Log In As Teacher
    Go To    ${BASE_URL}/pincode-verification.php
    Input Password    name:pincode    ${TEACHER_PIN}
    Click Button    css:button[type="submit"]
    Wait Until Location Contains    enroll.php
    Page Should Contain Element    name:course

TC06 Teacher Profile Displays Registration Number
    Log In As Teacher
    Go To    ${BASE_URL}/my-profile.php
    Textfield Value Should Be    name:studentregno    ${TEACHER_ID}

TC07 Enrollment History Loads
    Log In As Teacher
    Go To    ${BASE_URL}/enroll-history.php
    Page Should Contain    Enroll History

*** Keywords ***
Open Test Browser
    Open Browser    ${BASE_URL}/index.php    ${BROWSER}    options=add_argument("--headless=new");add_argument("--disable-gpu");add_argument("--no-sandbox");add_argument("--disable-dev-shm-usage")
    Set Window Size    1365    900

Log In As Teacher
    Go To    ${BASE_URL}/logout.php
    Go To    ${BASE_URL}/index.php
    Click Element    id:menu-student-signin
    Input Text    name:regno    ${TEACHER_ID}
    Input Password    name:password    ${TEACHER_PASSWORD}
    Click Button    css:button[type="submit"]
    Wait Until Location Contains    change-password.php
