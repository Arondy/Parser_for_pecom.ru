# Parser_pecom.ru
I created this parser for the website pecom.ru. It collects the data from xlsx file and searches the needed cities, then gets the data about price for different options and writes it in another xlsx file.
You'll need selenium and pandas libraries for this code to work.
It can be modified to get not only addresses, but also other values exactly from the xlsx file. Also, it won't write anything in the file if its work is interrupted by something in the website (e.g. if the city doesn't have a delivery, the code will be stopped).
