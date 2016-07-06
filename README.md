# servoyMailMerge

Small utility built using Servoy which allows one to repeatedly replace markers in a Microsoft Word Template or Outlook File Template.

# How to install
Apache POI libraries are used to parse the documents and are required.
Move the included POI folder to the Servoy developer's plugins directory prior to usage. Then import the dotxandoftToHTML.servoy solution within the Servoy IDE.

#Usage
To start the utility, execute the convert function.
The function takes in a pararmeter textToMerge. 
Which might look something like this:
```javascript
var textToMerge = {
	 '«FULLNAME»': 'Tuan Nguyen',
     '«TITLE»': 'Software Developer',
     '«COMPANY»': 'Servoy'
}
```
By default once convert has finished execution the file will be opened in the default browser for a quick view.

#Image support
There is rudimentary support for images in this build. 

#Filetypes
At this time it only supports .dotx (Microsoft Word Template) && .oft (Outlook File Template).