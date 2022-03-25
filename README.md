# Research for how many words per average page

This was made in Python because it was the language I had open at the time, and I made it just because I was interested to see how many words would fit on an average page. In the end I was able to do some data analysis, which was interesting to me. I was able to find how many words fit in a cm^2 on average, as well as the average amount of characters/page and obviously amount of words/page.


## Findings

In `Sheet2` of the Excel sheet, my findings are visible.

- There are on average 379.7623762 or 380 words per page on a LaTeX-generated PDF page, or 390.762376 (391) in a Word document;
- It took 0.338555928 seconds for each iteration of adding a new word, saving, converting to PDF, saving the data to Excel. This is about 22/65ths of a second, or close 1/3rd;
- There are on average 3518.019802 characters per LaTeX document this is about 3619.92094 characters for a Word document;
- I used a total of 494 words of the dictionary (which is not too good, but I only added 1 to the starting word each time a second page was found, I could have added the maximum amount of words to use more of the dictionary), this is roughly 2%;
- There are on average 1.085786104 words per cm^2, or about 1.1 words per cm^2. This corresponds to about 10.05843984 or 10.1 characters per cm^2;
- A word consists of on average 9.263739702 or 9.3 characters.


## Discussion

Because Word was acting up for me, and instead of being able to open a Word document as zip file, extracting the app.xml file and seeing the amount of pages it had, it just said I had 0 letters and 1 page, instead of trying to figure out why (which would be just about impossible), I decided to take the muuuuuuuch slower route of just converting the Word document to PDF and checking to see how many pages that had. Because I was using LaTeX to convert to PDF using Pandoc, about 11 words more fit on a Word document than on a PDF page in my case. This meant that doing the average amount of words I found plus 11 would get you about the amount of average words per page in Word.

I in no way ran this for long enough, I spent a day or 2 running it, in the beginning I planned on running it for about a week on my server, but because I didn't feel like upgrading Python and apparently string concatenation only existed in Python 3.6 which I found weird. I originally did try and upgrade to Python 3.10 for this, but it somehow didn't have access to a SSL/TLS module, which didn't allow me to install any modules. This was a problem I've had and fixed before, but I couldn't remember what the fix was, so I just decided to run it on my normal laptop instead.

The program is in no way the most efficient or accurate, I could've done this in so many different languages, ways, frameworks, etc. but this was the easiest and fastest one to program. On average it took about a third of a second to do one loop of adding a word and testing to see if it would fill a page, which is pretty insanely long, and most of this time was lost due to the conversion of document format.


## Settings to reproduce

The average size of my Word-documents were set to letter, so useful size (theoretical size - margins) was (`21.59 - 3.18*2`) 15.3cm width and (`27.94 - 2.54*2`) 22.86cm with font Cambria and font-size 11. Margins were the Office 2003 default, which were 3.18x2.54cm.