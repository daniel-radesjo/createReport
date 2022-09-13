#!/bin/bash

if [ -n "${1}" ]; then
  PARAMETERS="-r \"${1}\" -t \"${2}\" -o \"${3}\" -n \"${4}\" -c \"${5}\" -w \"${6}\""
fi

echo "Get result from textfiles" && ./get_items.sh
echo "Get EXIF information from files" && ./exif.sh
echo "Generate thumbnails from files" && ./convert.sh
./createReport.py $PARAMETERS

#Cleanup
echo "Cleanup generated files"
if [ -d "exif" ]; then
  rm -r exif
fi

if [ -f "result.txt" ]; then
  rm result.txt
fi

if [ -f "items.txt" ]; then
  rm items.txt
fi

find . -type d -name "Converted" -exec rm -r '{}' 2>&1 > /dev/null \; 2>&1 > /dev/null
