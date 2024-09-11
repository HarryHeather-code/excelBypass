#!/bin/bash

# Specify the path to the folder containing the Excel files with removed protection
outputFolder="no_protection"
mkdir $outputFolder

# Convert to zip
for file in *.xlsx; do

    mv -- "$file" "${file%.xlsx}.zip"
    
done

# Initialise arrays to print success and fail
SUCCESS=()
FAIL=()

# Remove space delimiter
IFS=""

# Loop through each Excel file in the folder
for file in *.zip; do

	# Extract the filename without extension
    filename=$(basename -- "$file")
    filename="${filename%.*}"

    # Specify the path to the output folder for the extracted sheets
    sheetsFolderPath="ExtractedSheets"

    # Extract the sheets from the zip file and remove sheet protection
    unzip -qq "$file" -d "$sheetsFolderPath"

    if (grep -Fq "sheetProtection algorithm" $sheetsFolderPath"/xl/worksheets/sheet1.xml")
    then
        #echo "$file = FAILED - cannot be unprotected without password"
        FAIL+=("$filename")
        
    else
        for xmlFile in "$sheetsFolderPath"/xl/worksheets/sheet*.xml; do
            # Remove sheet protection from the sheet
                sed -i 's/<sheetProtection[^>]*>//g' $xmlFile
        done

        # Move the edited sheets back to the zip file
        cd $sheetsFolderPath && zip -r ../$outputFolder/"$filename"_noSheetProtection.zip * && cd ..
        SUCCESS+=("$filename")
    fi

    # Clean up
    rm -rf "$sheetsFolderPath"
done

echo -e "\nComplete!" ${#SUCCESS[@]} "success and" ${#FAIL[@]} "fail.\nSuccesses saved to no_protection.\n"

# Convert all zips back to xlsx
for file in $outputFolder/*.zip; do

    mv -- "$file" "${file%.zip}.xlsx"

done
for file in *.zip; do

    mv -- "$file" "${file%.zip}.xlsx"

done


# Print success and fails
for s in ${SUCCESS[@]}; do   
    echo "SUCCESS = $s"
done
for f in ${FAIL[@]}; do
    echo "FAIL = $f" 
done