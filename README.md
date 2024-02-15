# reativacaoSistemaConexao
Script to reactivate employess on external system.

The script reads an Excel file with all the information needed and automatic reactivate the employee in the website using Selenium. The website sometimes requires a captcha before login in so the script connect to a Chrome instance running on a specific port.

# Libraries
- Pandas: used to read/write the Excel file;
- Selenium: used to automate the form filling process;
- SMTPLIB: used to send and e-mail when the script finished running;
- Datetime, OS - Python local libraries used on folder management and getting current date;
