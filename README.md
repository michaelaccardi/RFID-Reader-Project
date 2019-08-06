# RFID-Reader-Project
This project is still under development, however, currently achieves its basic purpose.

The purpose of this project is to:
1. Take a keyboard input from a key board enumeration RFID card reader
2. Create/Append a current day .txt file with each RFID swipe, recording the code and date time stamp.
3. Convert that inputted code into a students name using a CSV file(decoder.csv) as a reference database
4. Cross reference the current date and name to mark a student present for that day in a xlsx file(class_attendance.xlsx)


Currently Working On:
1. Making the program end at 0900. [Complete]
2. Checking if name is already in .txt brefore appending, if it is don't append.  [Not Working]
3. Breaking the input every 30 seconds using a timer function, in order to check time to break if no one is swiping currently [Researching]


Please note that names are taken from the top 2019 names list and qcodes are random and are not associated to the class.

