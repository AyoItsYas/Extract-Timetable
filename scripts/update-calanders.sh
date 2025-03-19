#!/bin/bash

if [ $# -ne 1 ]; then
  echo "Usage: $0 <FILENAME>"
  exit 1
fi

file=$1

mkdir -p ./.tmp
mkdir -p ./calanders

i=1
while read -r line; do
  if [[ $line =~ [^/]*$ ]]; then
    ANCHOR=$(echo "$line" | cut -d' ' -f1)
    HREF=$(echo "$line" | cut -d' ' -f2)
    FILENAME=$(echo "$HREF" | rev | cut -d'/' -f1 | rev)
  fi

  echo "Downloading sheet @ '$HREF' ..."
  wget --cookies=on --keep-session-cookies "$HREF&download=1" -O "./.tmp/$FILENAME.xlsx" --quiet &

done < "$file"

wait
echo "All downloads completed."

i=1
while read line; do
  if [[ $line =~ [^/]*$ ]]; then
    ANCHOR=$(echo $line | cut -d' ' -f1)
    HREF=$(echo $line | cut -d' ' -f2)
    FILENAME=$(echo $HREF | rev | cut -d'/' -f1 | rev)
  fi

  python3 main.py "./.tmp/$FILENAME.xlsx" --output "%SUMMARY% - %WS_TITLE% -- ($FILENAME).ics" --output_folder "./calanders/" --anchor "$ANCHOR" &

  i=$((i+1))
done < "$file"

wait
echo "All conversions completed."
