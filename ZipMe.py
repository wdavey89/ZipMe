import os, time, datetime, zipfile, send2trash, ctypes

def main():
    ctypes.windll.kernel32.SetConsoleTitleW('ZipMeV2') # Set the title of the console window.
    print(' --- This is the automated Zip Program --- ')
    print('Author: Wayne Davey')
    dirInput = input('Enter directory: ') # Enter directory in which to look for files.
    try:
        directory = os.chdir(dirInput)
    except FileNotFoundError:
        print('Directory does not exist, please try again. ')
        main()
    directory = os.getcwd() # Assign the directory to the current working directory.
    path = os.getcwd()
    archivePath = os.path.join(path, 'Archive') # Create the archive path by join 'directory' with an 'Archive' sub folder.
    if not os.path.exists(archivePath): # If the archive path doesn't exist, create it.
        os.makedirs(archivePath)
    month = getMonth() # Call 'getMonth' function.
    year = getYear() # Call 'getYear' function.
    zipMe(directory, archivePath, month, year) # Call 'zipMe' function passing in 4 variables.
    useAgain()


# This function requires you to input the month in which to based the file search on.
def getMonth():
    month = 0
    while month == 0 or month > 12:
        try:
            month = int(input('Enter the month: '))
        except ValueError:
            print('You entered an invalid value. ')
    return month

def getYear(): # This function requires you to input the year in which to based the file search on.
    year = 0
    while len(str(year)) != 4:
        try:
            year = int(input('Enter the year: '))
        except ValueError:
            print('You entered an invalid error. ')
    return year

# This is the function that creates the relevant log files and zips up files.
                       
def zipMe(directory, archivePath, month, year):
    zipDate = str(month)+ str(year) + '.zip' # Create a zip folder with the name of the month and year you inserted earlier.
    currentDateTime = datetime.datetime.now().strftime('%d-%m-%Y %H-%M-%S') # Get the current date and time
    logFileName = currentDateTime + '- ' + str(month) + str(year) + 'LOG.txt' # Create a log file with the current date and time, month and year.
    logDir = os.path.join('C:\\Workshop\\ZIPLOGS', logFileName) # Join these two directories together, this is the location of the log file.
    zipToLocation = os.path.join(archivePath, zipDate) # This is the location of where the ZIP file will be stored.
    newZip = zipfile.ZipFile(zipToLocation, 'w') # Open the ZIP file for writing.
    logFile = open(logDir, 'w') # Open log file for writing.
    logFile.write('Log created at: ' + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')+ '\n') # Write the current date and time to the log file.
    logFile.write('Directory in use:' + archivePath + '\n') # Write the archive path the log file.
    logFile.write('Files that have been added to ZIP Folder for month: ' + str(month) + ' and year: ' + str(year) + ' are: \n') # Write the text to the log file.
    print('''Creating ZIP Folder now... \nChecking for valid files to add to ZIP Folder... ''') # Print this to user screen.
    for items in os.listdir('.'): # List the contents of the current working directory.
        if os.path.isfile(items): # If there are files in this directory, continue.
            dt = datetime.date.fromtimestamp(os.path.getmtime(items)) # Get the Date Modified date from the files in the directory.
            if dt.month == month and dt.year == year: # If there are files with the modified month and modified year that match what the user inserted earlier, continue.
                files = [items] # Write the matched modified dated files to a list.
                for found in files: 
                    newZip = zipfile.ZipFile(zipToLocation, 'a') # Open the created ZIP file. 
                    newZip.write(found, compress_type=zipfile.ZIP_DEFLATED) # Write each file to the ZIP file.
                    logFile = open(logDir, 'a') 
                    logFile.write(found + '\n') # Append the file name written to the ZIP file to the log
                    logFile.close()
                    newZip.close()
                    print(found) # Print the file name of each file that is being written to the ZIP file.
                    send2trash.send2trash(found) # Send the written file to the recycle bin from it's original location.

    checkZipFolder(month, year, zipDate, archivePath) # Call this method once out of the above For Loop.

def checkZipFolder(month, year, zipDate, archivePath):
    os.chdir(archivePath) # Change the current working directory to the Archive Directory.
    checkZip = zipfile.is_zipfile(zipDate) # Check if a ZIP folder exists matching variable 'zipDate'
    try:
        with zipfile.ZipFile(zipDate, 'r') as zf: # Open the ZIP file for reading purposes.
            zf = zipfile.ZipFile.namelist(zf) # Checks if any files exist in the ZIP folder that matched the month and year of the date modified inserted earlier.
            print('The above files were successfully added to the ZIP Folder: {} '.format(zipDate))
    except zipfile.BadZipFile: # Except clause if the ZIP file contained no files, in which case the user entered a month and year that matched zero files with that date modified month and year.
        print('No files were found with the month: {} or year: {}.'.format(month, year))
        main() # Send user back to Main function.
            
def useAgain(): # Ask the user if they want to use the program again.
    zipAgain = input('Would you like to use the program again? Yes / No : ').lower()
    if zipAgain == 'y' or zipAgain == 'yes': # If they do, call Main()
        main()
    else:
        try:
            input('Press any key to exit.. ') # If they chose 'No', ask the user to press and key to exit.
        except SyntaxError:
            pass


if __name__ == '__main__':
    main()
