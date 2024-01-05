#!/bin/bash

if [ $# -ne 1 ]; then
  echo "Usage: $0 <FILENAME>"
  exit 1
fi

file=$1

mkdir -p ./.tmp
mkdir -p ./calanders

curl -i -c "./.tmp/.cookie.txt" "$line"

i=1
while read line; do
  if [[ $line =~ [^/]*$ ]]; then
    ANCHOR=$(echo $line | cut -d' ' -f1)
    HREF=$(echo $line | cut -d' ' -f2)
    FILENAME=$(echo $HREF | rev | cut -d'/' -f1 | rev)
  fi

  wget --cookies=on --load-cookies "./.tmp/.cookie.txt" --keep-session-cookies "$HREF&download=1" -O "./.tmp/$FILENAME.xlsx"

  python3 main.py "./.tmp/$FILENAME.xlsx" --output "%SUMMARY% -- ($FILENAME).ics" --output_folder "./calanders/" --anchor "$ANCHOR"
  i=$((i+1))
  echo "$ANCHOR $HREF $FILENAME"
done < "$file"
