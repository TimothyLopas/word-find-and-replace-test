*** Settings ***
Documentation       Template robot main suite.

Library    RPA.Word.Application
Library    RPA.FileSystem
Library    DateTime
Task Setup              Open Application    ${TRUE}
Suite Teardown          Quit Application

*** Variables ***
${FILE}=    find and replace test template.docx
${NEW_FILE}=    find and replace test.docx

${TEXT_REPLACEMENT_FILE}=    replacement text.docx

*** Tasks ***
Minimal task
    ${new_file_exists}=    Does File Exist    ${NEW_FILE}
    IF    ${new_file_exists}
        Remove File    ${NEW_FILE}
    END

    Open File    ${FILE}
    Replace Text    <<Company Name>>    Robocorp Inc.
    ${today}=    Get Current Date    result_format=%Y-%m-%d
    Replace Text    <<Date>>    ${today}

    Open File    ${TEXT_REPLACEMENT_FILE}
    Find Text    This is my second paragraph
    Move to Line Start
    Select Paragraph    3
    Copy Selection To Clipboard
    Close Document
    
    # Upon closing the second Word document the below text is 
    # able to be found in the original Word document
    Find Text    <<Malesuada2>>
    Move To Line Start
    Select Current Paragraph
    # For some reason the Pasting From Clipboard, while the correct content, will only past the text at the top of the document
    Paste From Clipboard
    Save Document As    ${NEW_FILE}
    Close Document