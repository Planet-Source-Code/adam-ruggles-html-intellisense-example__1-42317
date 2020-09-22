This read me Explains the what the data text file expects

The 1st character is what it uses for the delimeter for the rest of the line.
The 2nd thing it expects is the imagelist icon number for the small icon
The 3rd thing it expects is the What gets displayed in the ListView
The 4th thing it expects is the output string
the 5th thing it expects is the cursor offset from the begining of the string (number or @)
the 6th thing it expects is the "remove amount" it will remove that number of strings before the output string

an Example is:
,1,A,A HREF=",@,0
in the above example:
"," is the delimeter
"A" is the displayed string in the list view
"A HREF=" is outputed when it is selected
the '@' is a special character that means put the cursor at the end of the output
and it will remove 0 characters before outputing "A HREF="

Another Example is:
%2%&NBSP;%&nbsp;%@%1
now the following is true
"%" is the delimeter
"&NBSP;" is the displayed string in the list view
"&nbsp;" is outputed when it is selected
the '@' is a special character that means put the cursor at the end of the output
and it will remove 1 characters before outputing "&nbsp;" so it will remove the < symbol
