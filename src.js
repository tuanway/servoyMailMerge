/**
 * Open DOTX || OFT files and convert to HTML
 * @param {JSEvent} event
 * @param {Object} textToMerge Object containing values for mail merge replacement
 * @properties={typeid:24,uuid:"14168D03-8E8A-4EE3-A667-8DE4149FF53C"}
 */
function convert(event, textToMerge) {
    //create some test merge fields
    if (textToMerge == null) {
        textToMerge = {
            '«FULLNAME»': 'Tuan Nguyen',
            '«TITLE»': 'Software Developer',
            '«COMPANYML»': 'Servoy',
            '&laquo;fullname&raquo;': 'Tuan Nguyen'
        }
    }

    /**
     * Convert DOTX format to HTML using Apache POI
     * @private
     * @param {Boolean} open show in browser after conversion.
     * @SuppressWarnings(wrongparameters)
     */
    function dotxToHTML(open) {
        //convert to docx format for parsing
        var pkg = Packages.org.apache.poi.openxml4j.opc.OPCPackage.open(g);
        pkg.replaceContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
        //save as TEMP file
        var out = new java.io.FileOutputStream(new java.io.File('temp.docx'))
        pkg.save(out);
        out.close()

        //parse using XWPF POI
        var t = new Packages.org.apache.poi.xwpf.usermodel.XWPFDocument(new Packages.java.io.FileInputStream(new java.io.File('temp.docx')))

        //get paragraphs
        var p = t.getParagraphs()

        //get images
        var images = t.getAllPictures();

        // get runs
        for (var i = 0; i < p.size(); i++) {
            var r = p.get(i)['getRuns']();

            //inject merge text
            for (var j = 0; j < r['size'](); j++) {
                var text = r['get'](j).getText(0);
                if (text !== null) {
                    if (!textToMerge[text]) continue;
                    r['get'](j).setText(textToMerge[text], 0);
                }
            }
        }

        //get images and store into final location
        var imgLocation = 'word/media/'
        for (i = 0; i < images.size(); i++) {
            r = images.get(i)['getData']();
            var imgName = images.get(i)['getFileName']()
            Packages.org.apache.commons.io.FileUtils.writeByteArrayToFile(new java.io.File(imgLocation + imgName), r)
        }

        //Write changes to TMP file
        t.write(new java.io.FileOutputStream(new java.io.File('temp.docx')))

        var input = new Packages.java.io.FileInputStream(new java.io.File('temp.docx'))
        var doc = new Packages.org.apache.poi.xwpf.usermodel.XWPFDocument(input)

        //specify where load images for HTML
        var options;
        //options = new Packages.org.apache.poi.xwpf.converter.xhtml.XHTMLOptions.create().URIResolver(new Packages.org.apache.poi.xwpf.converter.core.FileURIResolver(new java.io.File("word/media")));

        var output = java.io.ByteArrayOutputStream()

        //convert to HTML
        Packages.org.apache.poi.xwpf.converter.xhtml.XHTMLConverter.getInstance().convert(doc, output, options);

        //close input
        input.close();
        if (open) {
            //open HTML file after conversion.
            var arr = new java.lang.String(output);
            var temp = plugins.file.createFile('temp.html');
            temp.setBytes(arr.getBytes(), true)
            plugins.file.openFile(temp);

        }
        return output;
    }

    /**
     * Convert OFT format to HTML using Apache POI-HSMF
     * @param {Boolean} open show in browser after conversion.
     * @private
     * @SuppressWarnings(wrongparameters)
     */
    //
    function oftToHTML(open) {
        //if file is OFT - parse using HSMF.
        var msg = new Packages.org.apache.poi.hsmf.MAPIMessage(g);
        var output = msg.getHtmlBody()

        //merge text
        for (var x in textToMerge) {
            output = output.replace(x, textToMerge[x])
        }

        if (open) {
            //open HTML file after conversion.
            var arr = new java.lang.String(output);
            var temp = plugins.file.createFile('temp.html');
            temp.setBytes(arr.getBytes(), true)
            plugins.file.openFile(temp);

        }

        return output;
    }

    //Show file dialog - restrict to DOTX or OFT file formats
    /** @type {plugins.file.JSFile} */
    var f = plugins.file.showFileOpenDialog(1, null, false, ['dotx', 'oft'])
    if (!f) {
        return '';
    }

    var g = new Packages.java.io.FileInputStream(f);

    if (f.getName().indexOf('.dotx') != -1) {
        return dotxToHTML(true);
    } else {
        return oftToHTML(true);
    }
}
