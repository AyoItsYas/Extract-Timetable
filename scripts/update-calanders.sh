#!/bin/bash

if [ $# -ne 1 ]; then
  echo "Usage: $0 <filename>"
  exit 1
fi

file=$1

mkdir -p ./.tmp
mkdir -p ./calanders

curl -i -c "./.tmp/.cookie.txt" "$line"

i=1
while read line; do
  if [[ $line =~ [^/]*$ ]]; then
    filename=${BASH_REMATCH[0]}
  fi

  wget --cookies=on --load-cookies "./.tmp/.cookie.txt" --keep-session-cookies "$line&download=1" -O "./.tmp/$filename.xlsx"

  python3 main.py "./.tmp/$filename.xlsx" --output "%SUMMARY% -- ($filename).ics" --output_folder "./calanders/"
  i=$((i+1))
done < "$file"
