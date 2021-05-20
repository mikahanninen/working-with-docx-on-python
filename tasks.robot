*** Settings ***
Documentation     Template robot main suite.
Library           CustomWord.py
Library           CustomWordWithTemplate.py
Library           RPA.PDF
Suite Setup       Create Template PDF File

*** Variables ***
${RESOURCES}      %{ROBOT_ROOT}${/}resources

*** Keywords ***
Create Template PDF File
    HTML to PDF    <h1>template pdf</h1>    ${RESOURCES}${/}template.pdf

*** Tasks ***
Minimal task
    Log To Console    \nWorking with Word documents
    Create Doc    %{ROBOT_ARTIFACTS}${/}demo.docx    ${RESOURCES}${/}okta.png
    ${embed_pdf}=    Create Dictionary    src=${RESOURCES}${/}template.pdf
    ...    target=${RESOURCES}${/}pywinauto.pdf
    #Transform Pdf Into Ole Object    ${RESOURCES}${/}pywinauto.pdf
    Create Doc With Template    ${RESOURCES}${/}pdf_template.docx
    ...    %{ROBOT_ARTIFACTS}${/}generated.docx
    ...    ${embed_pdf}
    # Create Doc With Template    ${RESOURCES}${/}zip_template.docx
    # ...    %{ROBOT_ARTIFACTS}${/}generated.docx
    # ...    ${embed_pdf}
    Log    Done.
