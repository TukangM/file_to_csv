#https://github.com/TukangM/file_to_csv
#License TukangM/file_to_csv is licensed under the "Creative Commons Zero v1.0 Universal"

$inputfile = "C:\path\to\path" 
$outputfile = "X:\output\to\path\fileoutput.csv"
$filetype = "*"

echo please wait
Get-ChildItem -Path $inputfile $filetype -Recurse | Export-Csv -Path $outputfile -Encoding ASCII -NoTypeInformation
echo allready done!
pause