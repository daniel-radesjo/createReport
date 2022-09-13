#!/bin/bash

if [ -d "CSV" ] && [ ! -d "CSV/Converted" ]; then
	cd CSV
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.csv*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi


if [ -d "DOC" ] && [ ! -d "DOC/Converted" ]; then
	cd DOC
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.doc*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "DOCX" ] && [ ! -d "DOCX/Converted" ]; then	
	cd DOCX
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.doc*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "EML" ] && [ ! -d "EML/Converted" ]; then	
	cd EML
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.htm*" -type f -exec echo '{}' \; -exec xvfb-run wkhtmltopdf --load-error-handling ignore {} Converted/'{}'.pdf \; -exec sleep 1 \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
	exit 0
fi

if [ -d "EMLX" ] && [ ! -d "EMLX/Converted" ]; then	
	cd EMLX
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.htm*" -type f -exec xvfb-run wkhtmltopdf {} Converted/'{}'.pdf \; -exec sleep 1 \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "JPEG" ] && [ ! -d "JPEG/Converted" ]; then
	cd JPEG
	mkdir -p Converted/Img/JPEG
	find . -maxdepth 1 -iname "*.jpeg" -type f -exec convert -density 150 -alpha remove -trim -resize 1200\> '{}' -quality 100 'Converted/Img/{}' \;
	mv Converted/Img/JPEG/*.* Converted/Img/; rmdir Converted/Img/JPEG
	cd ..
fi

if [ -d "JPG" ] && [ ! -d "JPG/Converted" ]; then
	cd JPG
	mkdir -p Converted/Img/JPG
	find . -maxdepth 1 -iname "*.jpg" -type f -exec convert -density 150 -alpha remove -trim -resize 1200\> '{}' -quality 100 'Converted/Img/{}' \;
	mv Converted/Img/JPG/*.* Converted/Img/; rmdir Converted/Img/JPG
	cd ..
fi

if [ -d "HTM" ] && [ ! -d "HTM/Converted" ]; then	
	cd HTM
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.htm" -type f -exec echo '{}' \; -exec xvfb-run wkhtmltopdf --load-error-handling ignore {} Converted/'{}'.pdf \; -exec sleep 1 \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "HTML" ] && [ ! -d "HTML/Converted" ]; then	
	cd HTML
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.html" -type f -exec echo '{}' \; -exec xvfb-run wkhtmltopdf --load-error-handling ignore {} Converted/'{}'.pdf \; -exec sleep 1 \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "LNK" ] && [ ! -d "LNK/Converted" ]; then	
	cd LNK
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*lnk.htm" -type f -exec echo '{}' \; -exec xvfb-run wkhtmltopdf --load-error-handling ignore {} Converted/'{}'.pdf \; -exec sleep 1 \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "MSG" ] && [ ! -d "MSG/Converted" ]; then	
	cd MSG
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.htm*" -type f -exec echo '{}' \; -exec xvfb-run wkhtmltopdf --load-error-handling ignore {} Converted/'{}'.pdf \; -exec sleep 1 \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "NUMBERS" ] && [ ! -d "NUMBERS/Converted/Img" ]; then
	cd NUMBERS
	mkdir -p Converted/Img
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "ODS" ] && [ ! -d "ODS/Converted" ]; then
	cd ODS
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.ods*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "ODT" ] && [ ! -d "ODT/Converted" ]; then
	cd ODT
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.odt*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "PAGES" ] && [ ! -d "PAGES/Converted/Img" ]; then
	cd PAGES
	mkdir -p Converted/Img
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "PDF" ] && [ ! -d "PDF/Converted" ]; then
	cd PDF
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv *.jpg Converted/Img/
	cd ..
fi

if [ -d "PNG" ] && [ ! -d "PNG/Converted" ]; then
	cd PNG
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.png*" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv *.jpg Converted/Img/
	cd ..
fi

if [ -d "RTF" ] && [ ! -d "RTF/Converted" ]; then
	cd RTF
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.rtf*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "SIE" ] && [ ! -d "SIE/Converted" ]; then
	cd SIE
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.s?*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "TIF" ] && [ ! -d "TIF/Converted" ]; then
	cd TIF
	mkdir -p Converted
	find . -maxdepth 1 -iname "*.tif*" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv *.jpg Converted
	cd ..
fi

if [ -d "TIFF" ] && [ ! -d "TIFF/Converted" ]; then
	cd TIFF
	mkdir -p Converted
	find . -maxdepth 1 -iname "*.tif*" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv *.jpg Converted
	cd ..
fi

if [ -d "TXT" ] && [ ! -d "TXT/Converted" ]; then
	cd TXT
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.txt" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "XLS" ] && [ ! -d "XLS/Converted" ]; then
	cd XLS
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.xls*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "XLSM" ] && [ ! -d "XLSM/Converted" ]; then
	cd XLSM
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.xls*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "XLSX" ] && [ ! -d "XLSX/Converted" ]; then
	cd XLSX
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.xls*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi

if [ -d "XLT" ] && [ ! -d "XLT/Converted" ]; then
	cd XLT
	mkdir -p Converted/Img
	find . -maxdepth 1 -iname "*.xlt*" -type f -exec soffice --headless --invisible --convert-to pdf {} --outdir Converted/ \;
	find Converted -maxdepth 1 -iname "*.pdf" -type f -exec convert -crop 1x1@ -density 150 -alpha remove -trim '{}'[0] -quality 100 '{}'.jpg \;
	mv Converted/*.jpg Converted/Img/
	cd ..
fi
