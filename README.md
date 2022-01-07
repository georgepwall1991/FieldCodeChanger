# FieldCodeChanger

I would like the program to read a .docx and search for a specific sequence of field codes eg.

{ FIELDCODE1 }/{ FIELDCODE2 }/{ FIELDCODE3 }

or

{ FIELDCODE1 }{ FIELDCODE2 }{ FIELDCODE3 }

or

{ FIELDCODE1 \*charformat }/{ FIELDCODE2  \*arabic }/{ FIELDCODE3 \*arabic }

or

{ FIELDCODE1 \*charformat }{ FIELDCODE2  \*arabic }{ FIELDCODE3 \*arabic }

or

{ FIELDCODE1 \*charformat }/{ FIELDCODE2  \*arabic \*charformat }/{ FIELDCODE3 \*arabic \*charformat }

or

{ FIELDCODE1 \*charformat }{ FIELDCODE2  \*arabic \*charformat }{ FIELDCODE3 \*arabic \*charformat }





Then edit the file so the document then reads:

{ FIELDCODE2 \*arabic \*charformat }/{ FIELDCODE3 \*arabic \*charformat }/{ FIELDCODE1 \*charformat }


Secondary Functionality (pretty essential):

Make the program able to apply the changes to every document in a folder.

Tertiary Functionality (nice to have but far from essential):

Update two spreadsheets of a specific format to indicate the documents that have been changed successfully by the program, and to highlight any documents that did not contain the target sequence (This would help identify improvements to the Target sequence to catch any missed cases.).
