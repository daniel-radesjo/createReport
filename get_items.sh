#!/bin/bash
#Extract information from filelist for extracted files in current directory
find . -iname "*]*.*" -type f | rev | cut -d"[" -f1 | rev | cut -d"]" -f1 | sed "s/.*/^&\t/" > items.txt
head -n1 *_FileList.txt > result.txt
grep -f items.txt *_FileList.txt >> result.txt

