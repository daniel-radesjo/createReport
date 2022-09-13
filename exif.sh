#!/bin/bash
#Run exiftool on every file in folders

mkdir -p exif
find ./* -type d -print0 | while read -d $'\0' d; do
  find "${d}" -type f -print0 | while read -d $'\0' f; do
    filename=$(echo "${f}" | cut -d"/" -f3)
    
    if [ ${#filename} -gt 0 ]; then
      exiftool -S "${f}" > "exif/${filename}.exif.txt"
    fi
  done
done
