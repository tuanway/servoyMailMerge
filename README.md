# servoyMailMerge

Small utility for use with mail merges in Servoy Smart Client
Apache POI is required and is included in the 'POI' folder. 

# Import solution
Move POI folder to the Servoy developer's plugins directory prior to usage.
To test funcionality import 'dotxandoftToHTML.servoy' solution.

#Usage
The convert function takes in one parameter, an object textToMerge.
It might look something like this:

var textToMerge = {
	 '«FULLNAME»': 'Tuan Nguyen',
     '«TITLE»': 'Software Developer',
     '«COMPANYML»': 'Servoy',
}

it will look for any mail merge fields in a DOTX/OFT file and replace them with the chosen value.