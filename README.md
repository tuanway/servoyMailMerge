#servoyMailMerge

Small utility built using Servoy which allows one to repeatedly replace markers in a Microsoft Word Template or Outlook File Template and convert the content into HTML for use with emails.

#How to install
Apache POI libraries are used to parse the documents and are required.
Add the neccessary libraries to the servoy classpath, then import the mailMerge.servoy solution within the Servoy IDE.

#Apache POI
Current releases can be found [here](http://archive.apache.org/dist/poi/release/bin/).
The lib folder also contains the version used in this build.

#Usage
If using the solution provided, click the 'select file' button to start the utility.  There is also an option to display converted html content in a browser post execution which is on by default.

The convert function is where most of the logic takes place.  Pass a pararmeter textToMerge with the desired markers to replace.
It might look something like this:
```javascript
var textToMerge = {
	 '«FULLNAME»': 'Tuan Nguyen',
     '«TITLE»': 'Software Developer',
     '«COMPANY»': 'Servoy'
}

convert(event,textToMerge,true);

```
Calling the convert command above will replace all markers in a selected document with those found in the textToMerge object and open the converted content in a browser after execution.


#Image support
There is rudimentary support for images in this build. 

#Filetypes
At this time it only supports .dotx (Microsoft Word Template) && .oft (Outlook File Template).
