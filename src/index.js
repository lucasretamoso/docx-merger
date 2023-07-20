var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

var Style = require('./merge-styles');
var Media = require('./merge-media');
var RelContentType = require('./merge-relations-and-content-type');
var bulletsNumbering = require('./merge-bullets-numberings');

function DocxMerger(options, files) {

    this._body = [];
    this._header = [];
    this._footer = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    this._files = [];
    var self = this;
    (files || []).forEach(function(file) {
        self._files.push(new JSZip(file));
    });
    this._contentTypes = {};

    this._media = {};
    this._rel = {};

    this._builder = this._body;

    this.insertPageBreak = function() {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        this._builder.push(pb);
    };

    this.insertRaw = function(xml) {

        this._builder.push(xml);
    };

    this.mergeBody = function(files) {

        var self = this;
        this._builder = this._body;

        RelContentType.mergeContentTypes(files, this._contentTypes);
        Media.prepareMediaFiles(files, this._media);
        RelContentType.mergeRelations(files, this._rel);

        bulletsNumbering.prepareNumbering(files);
        bulletsNumbering.mergeNumbering(files, this._numbering);

        Style.prepareStyles(files, this._style);
        Style.mergeStyles(files, this._style);

        var numLength = '<w:numId w:val='.length;

        files.forEach(function(zip, index) {
            //var zip = new JSZip(file);
            var xml = zip.file("word/document.xml").asText();
            xml = xml.substring(xml.indexOf("<w:body>") + 8);
            xml = xml.substring(0, xml.indexOf("</w:body>"));
            xml = xml.trim();
            var numId = xml.indexOf('<w:numId w:val=');
            while (numId !== -1) {
              var xmlAux = `${xml.substring(0, numId + numLength + 2)}${index.toString()}`;
              xml = `${xmlAux}${xml.substring(numId + numLength + 2)}`; 
              numId = xml.indexOf('<w:numId w:val=', numId + 1);
            }
            if (xml.lastIndexOf("<w:sectPr") === 0) {
                let tag = "</w:sectPr>";
                xml = xml.substring(xml.lastIndexOf(tag) + tag.length);
            } else {
                xml = xml.substring(0, xml.lastIndexOf("<w:sectPr"));
            }

            self.insertRaw(xml);
            if (self._pageBreak && index < files.length-1)
                self.insertPageBreak();
        });
    };

    this.save = function(type, callback) {
        var zip = this._files[0];

        var xml = zip.file("word/document.xml").asText();
        var style = '';
        var totalTags = [];
        this._files.forEach((file) => {
            var xmlBinFile = file.file("word/document.xml");
            if (xmlBinFile) {
                var xmlFile = xmlBinFile.asText();
                var indexFinishStyleDocument = xmlFile.indexOf("<w:body>");
                xmlFile = xmlFile.slice(
                xmlFile.indexOf("<w:document") + "<w:document".length,
                indexFinishStyleDocument
                );
                var ignorable = xmlFile.indexOf("mc:Ignorable");
                var ignorableText = "";
                if (ignorable !== -1) {
                var firstIndexQuotationMarks = xmlFile.indexOf("\"", ignorable);
                var secondIndexQuotationMarks = xmlFile.indexOf("\"", firstIndexQuotationMarks + 1);
                ignorableText = xmlFile.slice(ignorable, secondIndexQuotationMarks + 1);
                xmlFile = `${xmlFile.slice(0, ignorable).trim()}${xmlFile.slice(secondIndexQuotationMarks + 1).trim()}`
                }
                xmlFile = xmlFile.trim().replace(/ /g, "=");
                var tags = xmlFile.split("=");
                tags.forEach((tag, index) => {
                if (index % 2 === 0 && tags[index + 1] && tag) {
                    if (!totalTags.includes(tag)) {
                        style = `${style} ${tag}=${tags[index + 1].replace(/>/g, "")}`;
                        totalTags.push(tag);
                    }
                }
                });
                if (!style.includes("mc:Ignorable")) {
                    style = `${style} ${ignorableText}`
                }
            }
        });
        xml = `${xml.slice(
            0,
            xml.indexOf("<w:document")
        )}<w:document ${style}>${xml.slice(xml.indexOf("<w:body>"))}`;

        var startIndex = xml.indexOf("<w:body>") + 8;
        var endIndex = xml.lastIndexOf("<w:sectPr");

        xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(''));

        RelContentType.generateContentTypes(zip, this._contentTypes);
        Media.copyMediaFiles(zip, this._media, this._files);
        RelContentType.generateRelations(zip, this._rel);
        bulletsNumbering.generateNumbering(zip, this._numbering, this._files);
        Style.generateStyles(zip, this._style);

        zip.file("word/document.xml", xml);

        callback(zip.generate({ 
            type: type,
            compression: "DEFLATE",
            compressionOptions: {
                level: 4
            }
        }));
    };


    if (this._files.length > 0) {
        this.mergeBody(this._files);
    }
}


module.exports = DocxMerger;
