CompactAllMDBs is a Microsoft Access 2000/2002 (XP) application that will 

(1) Search for all specified files (default is .MDB and .MDA files) in a specified directory (and optionally, in all subdirectories below it)

(2) Display all such files found in a listbox with their sizes

(3) Compact all the files it found or only the ones selected from the list by the user

(4) Search for the same files it found the first time and compare their size as compacted with their original size, to see how much disk space has been reclaimed

It's handy to keep your ever-inflating MDB files of a manageable size, and it works nicely on networks. This can't really eb smoothly ported to MS Access 97 because it uses ENUMs in the code, which isn't supported by VBA in Office 97, to my understanding. 

Also included:

(1) A simple routine for calculating and displaying elapsed time for the process down to milliseconds.

(2) A "Browse For Folder" routine, based on a Windows API call (I originally wanted to make this start with a specified default folder, but I haven't had a chance to figure out how to do that yet)

(3) Uses the Access (2000/2002) routine to Compact and ReCompile the application when it's closed down.


In your VBA code, this application requires a Reference to MS DAO 3.6 library

If you have any comments or questions, drop a note to brianb@battleszone.com

Have an exceptional day!

  --Brian
    Middletown, CT USA
    http://www.BattlesZone.com/