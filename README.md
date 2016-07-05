# servoyMailMerge

Small utility for use with mail merges in Servoy Smart Client
Apache POI is required and is included in the 'POI' folder. 

# Import solution
Move POI folder to the Servoy developer's plugins directory prior to usage.
To test funcionality import 'dotxandoftToHTML.servoy' solution.

#Usage
The convert function takes in a pararmeter textToMerge.
It might look something like this:
```json
var textToMerge = {
	 '«FULLNAME»': 'Tuan Nguyen',
     '«TITLE»': 'Software Developer',
     '«COMPANYML»': 'Servoy',
}
```

The utility will look for any mail merge fields in selected file and replace them with the chosen value.  User has the option to return just the string or open as an html file.

#Images
There is rudimentary support for images in this build.  Has been tested only with the DOTX format. 

#Filetypes
Currently this tool only supports DOTX(word template) && OFT (outlook templates).