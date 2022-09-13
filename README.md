# createReport
Create a report from files with metadata and thumbnail of first page

# Required packages
- python-docx 0.8.11 (https://pypi.org/project/python-docx/)
- ImageMagick 7.1.0-48
- LibreOffce 7.4.0
- wkhtmltopdf 0.12.6 (https://wkhtmltopdf.org/)
- exiftool 12.44

# Create report
```
./createReport.sh

./createReport.sh result.txt template_en.docx report.docx Case1234 Client Worker
```
