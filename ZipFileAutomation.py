import sys, os, time, datetime, zipfile, send2trash, ctypes

def main():
    ctypes.windll.kernel32.SetConsoleTitleW('ZipFile Automation')
    if len(sys.argv) < 3:
        print('Not enough arguments were passed. Program changing to input mode.')
        getInfo()
    else:
        directoryInput = sys.argv[1]
        month = int(sys.argv[2])
        year = int(sys.argv[3])
        args(directoryInput, month, year)
        
# Asses the arguments passed and create the directory
def args(directoryInput, month, year):
    try:
        directory = os.chdir(directoryInput)
    except FileNotFoundError:
        print('Directory does not exit.')
        sys.exit()        
    directory = os.getcwd()
    archivePath = os.path.join(directory, 'Archive')
    if not os.path.exists(archivePath):
        os.makedirs(archivePath)
    zipMe(directory, archivePath, month, year)
    

# If no arguments are passed, user will be asked to input the directory, month & year
def getInfo():
    directory = input('Enter Directory: ')
    try:
        directory = os.chdir(directory)
    except FileNotFoundError:
        print('Directory does not exist, please try again. ')
        getInfo()
    directory = os.getcwd()
    archivePath = os.path.join(directory, 'Archive')
    if not os.path.exists(archivePath):
        os.makedirs(archivePath)
    month = 0
    while month == 0 or month > 12:
        try:
            month = int(input('Enter the month: '))
        except ValueError:
            print('You entered an invalid value. ')
    year = 0
    while len(str(year)) != 4:
        try:
            year = int(input('Enter the year: '))
        except ValueError:
            print('You entered an invalid error. ')
    zipMe(directory, archivePath, month, year)
    useAgain()
    
# This is where the logfile and ZipFile is created and written to.
def zipMe(directory, archivePath, month, year):
    zipDate = str(month)+ str(year) + '.zip'
    currentDateTime = datetime.datetime.now().strftime('%d-%m-%Y %H-%M-%S')
    logFileName = currentDateTime + ' - ' + str(month) + str(year) + ' LOG.txt'
    logFileDirectory = ('C:\\Workshop\\ZIPLOGS')
    if not os.path.exists(logFileDirectory):
        os.makedirs(logFileDirectory)
    logDir = os.path.join('C:\\Workshop\\ZIPLOGS', logFileName)
    zipToLocation = os.path.join(archivePath, zipDate)   
    newZip = zipfile.ZipFile(zipToLocation, 'w')
    logFile = open(logDir, 'w')
    logFile.write('Log created at: ' + datetime.datetime.now().strftime('%d-%m-%Y %H:%M:%S')+ '\n')
    logFile.write('Directory in use: ' + archivePath + '\n')
    logFile.write('Files that have been added to ZIP Folder for month: ' + str(month) + ' and year: ' + str(year) + ' are: \n')
    print('''Creating ZIP Folder now... \nChecking for valid files to add to ZIP Folder... ''')
    for items in os.listdir('.'):
        if os.path.isfile(items):
            dt = datetime.date.fromtimestamp(os.path.getmtime(items))
            if dt.month == month and dt.year == year:
                files = [items]
                for found in files:
                    newZip = zipfile.ZipFile(zipToLocation, 'a')
                    newZip.write(found, compress_type=zipfile.ZIP_DEFLATED)
                    logFile = open(logDir, 'a')
                    logFile.write(found + '\n')
                    logFile.close()
                    newZip.close()
                    print(found)
                    os.unlink(found)
    checkZipFolder(month, year, zipDate, archivePath)
    
# Checks to see if a valid zip folder was created and if it contains files. 
def checkZipFolder(month, year, zipDate, archivePath):
    os.chdir(archivePath)
    checkZip = zipfile.is_zipfile(zipDate)
    try:
        with zipfile.ZipFile(zipDate, 'r') as zf:
            zf = zipfile.ZipFile.namelist(zf)
            print('The above files were successfully added to the ZIP Folder: {} '.format(zipDate))
    except zipfile.BadZipFile:
        print('No files were found with the month: {} or year: {}'.format(month, year))
        main()
        
# Asks the user if they want to use the program again. You'll only get this if you've inputted information manually.
def useAgain():
    zipAgain = input('Would you like to use the program again? Yes / No : ').lower()
    if zipAgain == 'y' or zipAgain == 'yes':
        main()
    else:
        try:
            input('Press any key to exit.. ')
        except SyntaxError:
            pass



if __name__ == '__main__':
    main()
