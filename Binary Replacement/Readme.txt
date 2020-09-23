Readme for Binary String Replacer
=================================

This program works by simply using a binary file mode to read data
from any file into a buffer of changeable size, and then searching the
buffer for specified text and replacing any matches with desired text.
It should be noted that this will not work on some files. If you use
a hexadecimal editing program to view some .exe files, you will notice
that some .exe files store text in the following format:
	H.e.l.l.o...w.o.r.l.d.
while others store text in a predictable format:
	Hello world
This program will only work on files that store text with the second
method. If the string you want to be replaced is not properly replaced,
try messing around with the buffer size.
