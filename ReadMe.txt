OCR Project


There is an included recognition file (LetterInfo.txt) which is capital letters A-Z that I have already trained.

You may have to retrain the app for your own handwriting styles.

To train, Draw a letter...hit Recognize, and when it is done analyzing the image, and possibly making a guess..press the Train button to tell it what the letter should be.

Enabling AutoTrain will simply force the program to ask you what the letter is, if it has no guess to make.

Enabling Debug shows extra windows which help you see what the program is looking at it, and why it may have made another match.
This slows it down a bit since it has to draw in those extra boxes.

Press Reset after recognizing before attempting another recognition.

You can draw by clicking in the left picturebox, and can erase by doing the same with the right mouse button.

Erasing will also instantly erase any gridlines that are drawn there from the previous recognition.

#####
Sampling rate - Rate at which pixels are read on both the X and Y axis...default is 2,2 which reads every 2nd pixel on both axis', and seems to work the best.  If you change this value the current LetterInfo.txt file will probably be of no use to you, and may even give you bad matches. (If you change Sample rate, make yourself a new letterinfo.txt file, and backup the old one)

Tolerance Level - Max number of pixels a match can be off and still be considered a match. Default is 30 which seems to work the best.

Max # of bad sectors - Specifies how many "sectors" are allowed to be out of the max tolerance level. This was a test, and anything more than 0 gave horrible results.
#####

Any improvements or suggestions would happily be accepted when emailed to ntamayo@tampabay.rr.com

Thanks,
Nelson
ntamayo@tampabay.rr.com