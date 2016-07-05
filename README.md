# servoyMailMerge

Small utility for use with mail merges in a Servoy Smart Client solution.
This was created using Apache POI and it is required for working functionality.

# Import solution
Move the included POI folder to the Servoy developer's plugins directory prior to usage.
To test funcionality import 'dotxandoftToHTML.servoy' solution.

#Usage
The convert function takes in a pararmeter textToMerge.
It might look something like this:
```javascript
var textToMerge = {
	 '«FULLNAME»': 'Tuan Nguyen',
     '«TITLE»': 'Software Developer',
     '«COMPANY»': 'Servoy'
}
```

The utility will look for any mail merge fields in selected file and replace them with the chosen value.  User also has the option to open the html file in a browser for quick view.

#Images
There is rudimentary support for images in this build.

#Filetypes
Currently this tool only supports .dotx (Microsoft Word Template) && .oft (Outlook File Template).