# GUI-Based-Marksheet-Generator

## Getting Started

### REQUIREMENTS
Python (3.9.6) must be installed in the system.

### Instructions to Run

#### CLONE
Clone the repo on your system by using `git clone https://github.com/abhishekgupta3/GUI-Based-Marksheet-Generator.git`.

Setup project environment with [virtualenv](https://virtualenv.pypa.io) and [pip](https://pip.pypa.io).

```bash
$ virtualenv env
$ source env/bin/activate (for LINUX)
$ env\Scripts\activate (for WINDOWS)
$ pip install -r requirements.txt
```

In marksheetGenerator/settings.py

`Update EMAIL_HOST_USER as your email address`

Create a .env file in the same folder

Add this line in the .env file

`EMAIL_PASS={YOUR_EMAIL_PASSWORD}`

#### Run the server

```bash
$ python manage.py runserver
```
Open `https://localhost:8000/` in the browser