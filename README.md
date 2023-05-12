# HappyBirthday
## A Python based application that sends emails to clients on their birthday
![birthday-card](https://github.com/jnordst/HappyBirthday/assets/12515630/2e33c36a-9eb4-4f7e-9929-d6e3fad73bd3)

## How it Works
- Iterates through an Excel file containing client information
- Checks the birthday column for each row
- If their birthday is today, sends client a personalized email from your organization
- Uses validation to ensure the email is valid before sending an email
- Obscures email credentials within environment variables for security

## Sample Email
- Automatically uses information from the spreadsheet to fill in information
- First Name / Last Name / Last Visit Date 
![image](https://github.com/jnordst/HappyBirthday/assets/12515630/b955535b-14d7-494e-ba07-405c327bc5ec)

Jacob Nordstrom | April 2023
