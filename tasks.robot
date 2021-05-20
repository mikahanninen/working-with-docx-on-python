*** Settings ***
Documentation     Template robot main suite.
Library           CustomWord.py
Library           CustomWordWithTemplate.py

*** Variables ***
${RESOURCES}      %{ROBOT_ROOT}${/}resources

*** Tasks ***
Minimal task
    Create Doc    %{ROBOT_ARTIFACTS}${/}demo.docx    ${RESOURCES}${/}okta.png
    Create Doc With Template    ${RESOURCES}${/}pdf_template.docx
    ...    %{ROBOT_ARTIFACTS}${/}generated.docx
    ...    ${RESOURCES}${/}pywinauto.pdf
    Log    Done.
