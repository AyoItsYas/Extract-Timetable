#!/bin/bash

if [ $# -ne 1 ]; then
  echo "Usage: $0 <filename>"
  exit 1
fi

file=$1

curl -i -c "./data/.cookie.txt" "$line"

i=1
while read line; do
  wget --cookies=on --load-cookies "./data/.cookie.txt" --keep-session-cookies "$line" -O "data/calander-$i.xlsx"

  python3 main.py "data/calander-$i.xlsx"
  i=$((i+1))
done < $file