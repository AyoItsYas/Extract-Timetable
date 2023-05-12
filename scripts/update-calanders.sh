#!/bin/bash

if [ $# -ne 1 ]; then
  echo "Usage: $0 <filename>"
  exit 1
fi

file=$1

mkdir -p ./.tmp
mkdir -p ./calanders

curl -i -c "./.tmp/.cookie.txt" "$line&download=1"

i=1
while read line; do
  wget --cookies=on --load-cookies "./.tmp/.cookie.txt" --keep-session-cookies "$line&download=1" -O "./.tmp/calander-$i.xlsx"

  python3 main.py "./.tmp/calander-$i.xlsx" --output_folder "./calanders/"
  i=$((i+1))
done < $file