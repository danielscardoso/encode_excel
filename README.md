I had a very big excel database which resulted from the server output of page hits on each webpage of a given website per date. 
There were various character encoding errors in the URL strings, and various encodings used simultaneously in the same file. 
This script allows you to correct the vast majority of those enconding errors by first detecting the encoding of each URL and then replacing it with the proper characters.
It also has some hand-defined replacement operations for fringe cases that I was not able to automate.
It also visually signals to the user the cells in which it was unable to do so.
It was made with European Portuguese as the use case

The script assumes that the file to be changed is in the same directory as the script.

IMPORTANT:
For the script to work, you need to either edit the filename on line 107, or name the file you want it to run on "teste.xlsx"
It outputs the result into a new file, on the same directory.
