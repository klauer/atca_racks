#!/bin/bash

find "/Volumes/Groups/LCLS/LCLS_II/LCLS II SC/Racks/Rack_Profile_Masters_LCLS-II" |grep -iv -e old -e bcs -e future -e "~" |grep -i xls > rack_filenames.txt

while read fn; do
  echo "Copying $fn"
  cp -f "$fn" copied/
done <rack_filenames.txt
