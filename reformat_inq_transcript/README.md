# Reformat InqScribe transcript

Convert excel to inqscribe. Not sure how to do I/O streaming from cmd to parse in InqScribe. Might work on that.

Currently:
- Checking which Sheet is not empty to convert to df. Might cause error if more than one sheet has text
- Allows for the 4th column to be translation text
- Runs a bat file to pop up text and inqscribe to copy and paste
- First line in txt file is for save file name