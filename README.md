# Certificate-Generator
This application will automatically create &amp; mail the certificates using the Certificate image and the recipients details in an Excel sheet  

## Getting Started
### Prerequisites
You need to have Python installed on your PC. If you do not have you can [install from here](https://www.python.org/downloads/)

### Installation

1. Clone the repo
   ```sh
   git clone https://github.com/githubcrce/Certificate-Generator.git
   ```
3. Install the Requirements
   ```sh
   pip install -r requirements.txt
   ```
4. Create a `credentials.json` file
   ```sh
   {
    "email" : "YOUR EMAILID",
    "password" : "YOUR EMAILID PASSWORD",
    "smtp_server" : "smtp.gmail.com",
    "smtp_port" : 587,
    "email_subject" : "YOUR EMAILID SUBJECT",
    "email_body" : "YOUR EMAILID BODY",
    "students_sheet" : "students.xls",
    "picture_certificate_template" : "sample-template.png",
    "path_to_folder": "",
    "path_to_font":"YOUR FONT ttf file"
   }
   ```

6. Run the Certificate Generator
   ```sh
   python certificate_generator.py
   ```
