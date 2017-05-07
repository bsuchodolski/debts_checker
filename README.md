# debts_checker
Small script to help lawyer friend make her work less tedious

Script checks for few values in docx document for every form that resides in this document and compares these with xlsx spreadsheet, which is source of proper data.

Why not create all these forms with python from xlsx in the first place? Because the docx document is collection of motions submitted by individuals. The only required work is to check few key things.

Thing to do/check:
- add order number for every motion,
- check credit value and compare it with xlsx spreadsheet,
- check if correct creditor group is filled in.

![alt text](https://raw.githubusercontent.com/bsuchodolski/debts_checker/master/document.png)

All data - names of companies, invoice numbers, dates and values - in example docx and xlsx submitted here were anonymized.
