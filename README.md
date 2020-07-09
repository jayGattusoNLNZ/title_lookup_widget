## Description 

Runs a floating window that sits on top of your windows and waits for a title to be entererd in the text box. 

At runtime it looks for the latest version of the title_lookup reference sheet in the master store. If it can't see that, it looks for one thats in the root directory. The one included is from 1st July 2020, so might not contain all the titles. 

Start typing the title for the publication you're coding, and it will dynamically show you names that match

When you  see your title, double click on it, or -> cursor arrow. 

It will select, and show the MMS ID. 

The "Copy MMS to clipboard" button does exactly that. The MMS is now visiable to the windows clipboard. 

The "Get Rcv Note" button reaches out to ALMA to get the recieving note for that title

[note, the display is a bit messy, and not the same as ALMA could be tidied up] 

[note. you can tinker with the height and width vars on line #27 if its not working out]

[note. this call is super slow. It can take minutes to come back from ALMA]

[note. you need an ALMA API key for this to work, and it has to be in stored in a text file as pointed to in the script] 

The "Quit" button does exactly that. 

## Secrets file

Its set in code to be `secrets = r"c:/source/secrets"` on line #12. Change this to suit your environment. 

this is a plain old text file. Must contain the below, changed to reflect your ALMA API keys. 

    [configuration]
    PRODUCTION = my_production_key
    SANDBOX = my_sandbox_key
