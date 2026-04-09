# Classroom Assignment Downloader

Python project to extract student submissions and download attachments from a Google Classroom assignment.

## Why I built this
This project solves a practical workflow problem: collecting assignment deliverables and attached files from Google Classroom in a structured way.

It also reflects the kind of work I enjoy most: combining automation, APIs, file handling, and traceable outputs to reduce manual effort.

## Main features
- Connects to Google Classroom using OAuth
- Lists courses and assignments
- Retrieves student submissions
- Identifies attached Drive files
- Downloads files stored in Google Drive
- Exports Google Docs, Sheets, and Slides to usable formats
- Organizes outputs by course / assignment / student
- Generates a manifest report in CSV format

## Tech stack
- Python
- Google Classroom API
- Google Drive API
- Pandas
- OAuth 2.0

## Project structure
```text
classroom-downloader/
├── README.md
├── requirements.txt
├── .gitignore
├── .env.example
├── credentials/
├── data/
├── src/
└── tests/
