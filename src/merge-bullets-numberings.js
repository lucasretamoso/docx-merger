var XMLSerializer = require("xmldom").XMLSerializer;
var DOMParser = require("xmldom").DOMParser;

var prepareNumbering = function (files) {
  var serializer = new XMLSerializer();

  files.forEach(function (zip) {
    var xmlBin = zip.file("word/numbering.xml");
    if (!xmlBin) {
      return;
    }
    var xmlString = xmlBin.asText();
    var xml = new DOMParser().parseFromString(xmlString, "text/xml");
    var nodes = xml.getElementsByTagName("w:abstractNum");

    for (var node in nodes) {
      if (/^\d+$/.test(node) && nodes[node].getAttribute) {
        var absID = nodes[node].getAttribute("w:abstractNumId");
        nodes[node].setAttribute("w:abstractNumId", absID + index);
        var pStyles = nodes[node].getElementsByTagName("w:pStyle");
        for (var pStyle in pStyles) {
          if (pStyles[pStyle].getAttribute) {
            var pStyleId = pStyles[pStyle].getAttribute("w:val");
            pStyles[pStyle].setAttribute("w:val", pStyleId + "_" + index);
          }
        }
        var numStyleLinks = nodes[node].getElementsByTagName("w:numStyleLink");
        for (var numstyleLink in numStyleLinks) {
          if (numStyleLinks[numstyleLink].getAttribute) {
            var styleLinkId = numStyleLinks[numstyleLink].getAttribute("w:val");
            numStyleLinks[numstyleLink].setAttribute(
              "w:val",
              styleLinkId + "_" + index
            );
          }
        }

        var styleLinks = nodes[node].getElementsByTagName("w:styleLink");
        for (var styleLink in styleLinks) {
          if (styleLinks[styleLink].getAttribute) {
            var styleLinkId = styleLinks[styleLink].getAttribute("w:val");
            styleLinks[styleLink].setAttribute(
              "w:val",
              styleLinkId + "_" + index
            );
          }
        }
      }
    }

    var numNodes = xml.getElementsByTagName("w:num");

    for (var node in numNodes) {
      if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
        var ID = numNodes[node].getAttribute("w:numId");
        numNodes[node].setAttribute("w:numId", ID + index);
        var absrefID = numNodes[node].getElementsByTagName("w:abstractNumId");
        for (var i in absrefID) {
          if (absrefID[i].getAttribute) {
            var iId = absrefID[i].getAttribute("w:val");
            absrefID[i].setAttribute("w:val", iId + index);
          }
        }
      }
    }

    var startIndex = xmlString.indexOf("<w:numbering ");
    xmlString = xmlString.replace(
      xmlString.slice(startIndex),
      serializer.serializeToString(xml.documentElement)
    );

    zip.file("word/numbering.xml", xmlString);
    // console.log(nodes);
  });
};

var mergeNumbering = function (files, _numbering) {
  // this._builder = this._style;

  // console.log("MERGE__STYLES");

  files.forEach(function (zip) {
    var xmlBin = zip.file("word/numbering.xml");
    if (!xmlBin) {
      return;
    }
    var xml = xmlBin.asText();

    xml = xml.substring(
      xml.indexOf("<w:abstractNum "),
      xml.indexOf("</w:numbering")
    );

    _numbering.push(xml);
  });
};

var generateNumbering = function (zip, _numbering, files) {
  var xmlBin = zip.file("word/numbering.xml");
  if (!xmlBin) {
    return;
  }
  var xml = xmlBin.asText();
  var styleNumbering = xml.slice(
    xml.indexOf("<w:numbering") + "<w:numbering".length,
    xml.indexOf("</w:numbering>") === -1 ? xml.indexOf("/>") : xml.indexOf(">")
  );
  styleNumbering = styleNumbering.trim();

  files.forEach((file, index) => {
    var xmlBinFile = file.file("word/numbering.xml");
    if (xmlBinFile) {
      var xmlFile = xmlBinFile.asText();
      var indexFinishNumbering = xmlFile.indexOf("</w:numbering>");
      var indexAbstract = xmlFile.indexOf("<w:abstractNum");
      var indexNum = xmlFile.indexOf("<w:num");
      var finishIndexToXMLFile =
        indexFinishNumbering === -1
          ? xmlFile.indexOf("/>")
          : indexAbstract === -1
          ? indexNum === -1
            ? xmlFile.lastIndexOf(">")
            : indexNum
          : indexAbstract;
      xmlFile = xmlFile.slice(
        xmlFile.indexOf("<w:numbering") + "<w:numbering".length,
        finishIndexToXMLFile
      );
      xmlFile = xmlFile.replace(/ /g, "=");
      var tags = xmlFile.split("=");
      tags.forEach((tag, index) => {
        if (index % 2 !== 0) {
          if (!styleNumbering.includes(tag)) {
            styleNumbering = `${styleNumbering} ${tag}=${tags[index + 1]}`;
          }
        }
      });
    }
  });
  var xmlGenerated = `${xml.slice(
    0,
    xml.indexOf("<w:numbering")
  )}<w:numbering ${styleNumbering}>${_numbering.join("")}</w:numbering>`;

  zip.file("word/numbering.xml", xmlGenerated);
};

module.exports = {
  prepareNumbering: prepareNumbering,
  mergeNumbering: mergeNumbering,
  generateNumbering: generateNumbering,
};
