# excel-copy-network-files
Simple macro to copy list of networked files into a single directory

Note: this is currently designed to run on PCs

1. Open file in Excel
2. Adjust "directoryPath" variable to a writeable directory, or create folder called "scripts" in C:\
3. In the "Files" sheet, enter the full path of the file in column A starting at cell 2. You may add as many rows as you need.
4. Run the macro "Copy_Files" 
5. Files will be copied to the specified directory and highlighted in blue to designate a sucessful copy
