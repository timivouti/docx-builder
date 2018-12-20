var Docxtemplater = require('docxtemplater');
var JSZip = require('jszip');
var convert = require('xml-js');

var systemXmlRelIds = 
{	
	"styles.xml" : "rId1",
	"settings.xml" : "rId2",
	"webSettings.xml" : "rId3",
	"footnotes.xml" : "rId4",
	"endnotes.xml" : "rId5",
	"header1.xml" : "rId6",
	"footer1.xml" : "rId7",
	"fontTable.xml" : "rId8",
	"theme1.xml" : "rId9"
}

exports.Document = function() {
	
	this._body = [];
	this._header = [];
	this._footer = [];
	this._builder = this._body;
    this._bold = false;
	this._italic = false;
	this._underline = false;
	this._font = null;
	this._size = null;
	this._alignment = null;
	
	
	this.beginHeader = function() 
	{
		this._builder = this._header;
	}
	
	this.endHeader = function()
	{
		this._builder = this._body;
	}
	
	this.beginFooter = function() 
	{
		this._builder = this._footer;
	}
	
	this.endFooter = function()
	{
		this._builder = this._body;
	}
	
	this.setBold = function(){
		this._bold = true;
	}
	
	this.unsetBold = function(){
		this._bold = false;
	}
	
	this.setItalic = function(){
		
		this._italic = true;
	}
	
	this.unsetItalic = function(){
		
		this._italic = false;
	}
	
	this.setUnderline = function(){
		
		this._underline = true;
	}
	
	this.unsetUnderline = function(){
		
		this._underline = false;
	}
	
	this.setFont = function(font){
		this._font = font;
	}
	
	this.unsetFont = function() {
		this._font = null;
	}
	
	this.setSize = function(size){
		this._size = size;
	}
	
	this.unsetSize = function(){
		this._size = null;
	}
	
	this.rightAlign = function(){
		this._alignment = "right";
	}
	
	this.centerAlign = function(){
		this._alignment = "center";
	}
	
	this.leftAlign = function(){
		this._alignment = null;
	}
	
	this.insertPageBreak = function()
	{
		var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';
				  
		this._builder.push(pb);
    }

	this.beginTable = function(options){
		
		if(!options)
		{
			this._builder.push('<w:tbl>');
		}
		else
		{
			options = options || { borderSize: 4, borderColor: 'auto' };
			this._builder.push('<w:tbl><w:tblPr><w:tblBorders> \
				<w:top w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:left w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:bottom w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:right w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:insideH w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				<w:insideV w:val="single" w:space="0" w:color="' + options.borderColor + '" w:sz="' + options.borderSize + '"/> \
				</w:tblBorders>	\
			</w:tblPr>');
		}
	}
	
	this.insertRow = function(){
		
		this._builder.push('<w:tr><w:tc>');
	}
	
	this.nextColumn = function(){
		this._builder.push('</w:tc><w:tc>');
	}
	
	this.nextRow = function(){
		this._builder.push('</w:tc></w:tr><w:tr><w:tc>');
	}
	
	this.endTable = function(){
		this._builder.push('</w:tc></w:tr></w:tbl>');
	}
	
    this.insertText = function(text) {
		
		var p = '<w:p>' +
		
			(this._alignment ? ('<w:pPr><w:jc w:val="' + this._alignment + '"/></w:pPr>') : '') +
			
			'<w:r> \
				<w:rPr>' +
				
				    (this._size ? ('<w:sz w:val="' + this._size + '"/>') : "") +
					(this._bold ? '<w:b/>' : "") +
					(this._italic ? '<w:i/>' : "") +
					(this._underline ? '<w:u w:val="single"/>' : "") +
					(this._font ? ('<w:rFonts w:hAnsi="' + this._font + '" w:ascii="' + this._font + '"/>') : "")					
					
					+
				'</w:rPr> \
				<w:t>[CONTENT]</w:t> \
			</w:r> \
		</w:p>'
		
        this._builder.push(p.replace("[CONTENT]", text));
    }
	
	this.insertRaw = function(xml){
		
		this._builder.push(xml);
	}
	
	this._replaceRIds = function(xml, replacements)
	{
		var xmlBuilder = [];
	    var startingIndex = 0;
		for(var i=0; i < xml.length; i++)
		{
		    if(xml[i] == '"' && xml[i+1] == 'r' && xml[i+2] == 'I' && xml[i+3] == 'd')
		    {
				var oldRId = ["rId"];
				i = i+4;
				while(xml[i] != "\"")
				{
					oldRId.push(xml[i]);
				    i++;
				}
			   
				oldRId = oldRId.join("");
				var newRId = replacements[oldRId] || oldRId;
			   
				xmlBuilder.push('"');
				xmlBuilder.push(newRId);
				xmlBuilder.push('"');
			   
			}
			else 
				xmlBuilder.push(xml[i]);
	    }
	   
		return xmlBuilder.join("");
	}
	
	this._utf8ArrayToString = function(array) {
		var out, i, len, c;
		var char2, char3;

		out = "";
		len = array.length;
		i = 0;
		while(i < len) {
		c = array[i++];
		switch(c >> 4)
		{ 
		  case 0: case 1: case 2: case 3: case 4: case 5: case 6: case 7:
			// 0xxxxxxx
			out += String.fromCharCode(c);
			break;
		  case 12: case 13:
			// 110x xxxx   10xx xxxx
			char2 = array[i++];
			out += String.fromCharCode(((c & 0x1F) << 6) | (char2 & 0x3F));
			break;
		  case 14:
			// 1110 xxxx  10xx xxxx  10xx xxxx
			char2 = array[i++];
			char3 = array[i++];
			out += String.fromCharCode(((c & 0x0F) << 12) |
						   ((char2 & 0x3F) << 6) |
						   ((char3 & 0x3F) << 0));
			break;
		}
		}

		return out;
	}
	
	this.rels = [];
	
  this.getExternalDocxRawXml = function (docxData) {
    try {
      var zip = new JSZip(docxData);
    } catch (error) {
      console.log(error);
      return;
    }
    

    var xml = this._utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
    xml = xml.substring(xml.indexOf("<w:body>") + 8);
    xml = xml.substring(0, xml.indexOf("</w:body>"));

    var relsXml = this._utf8ArrayToString(zip.file("word/_rels/document.xml.rels")._data.getContent());
    var replacements = null;

    while (relsXml.indexOf("<Relationship") != -1) {
      relsXml = relsXml.substring(relsXml.indexOf("<Relationship") + 13);
      relsXml = relsXml.substring(relsXml.indexOf("Id=\"") + 4);
      var id = relsXml.substring(0, relsXml.indexOf("\""));
      relsXml = relsXml.substring(relsXml.indexOf("Type=\"") + 6);
      var type = relsXml.substring(0, relsXml.indexOf("\""));
      relsXml = relsXml.substring(relsXml.indexOf("Target=\"") + 8);
      var target = relsXml.substring(0, relsXml.indexOf("\""));

      var filename = target.indexOf("/") != -1 ? target.substring(target.lastIndexOf("/") + 1) : target;
      var zipPath = target.startsWith("../") ? target.substring(3) : ("word/" + target);


      var newId = systemXmlRelIds[filename];
      var newTarget = target;

      if (zip.file(zipPath) && zip.file(zipPath)._data != null) {
      if (!newId) {
        var hrtime = this.hrtime();
        var rand = hrtime[0] + "" + hrtime[1];
        newId = id + "_" + rand;
        newTarget = target.split('/');
        newTarget[newTarget.length - 1] = rand + "_" + newTarget[newTarget.length - 1];
        newTarget = newTarget.join('/');
      }

      this.rels.push({
        id: id,
        newId: newId,
        data: zip.file(zipPath)._data.getContent(),
        zipPath: zipPath,
        filename: filename,
        type: type,
        target: target,
        newTarget: newTarget
      });


      replacements = replacements || {};
      replacements[id] = newId;
    }


    if (replacements)
      xml = this._replaceRIds(xml, replacements);
  }
		
		return xml;
	}
	
	this.insertDocxSync = function(file){
		
      var xml = this.getExternalDocxRawXml(file);
		this.insertRaw(xml);
	}
	
	
	this.save = function(template, isDraft){
		
		var zip = new JSZip(template);
      var filesToSave = {};

    if (this.rels.length > 0)
		{
			var relsXmlBuilder = [];
			
			for(var i=0; i < this.rels.length; i++)
			{
				var rel = this.rels[i];
              var saveTo = rel.newTarget.startsWith("../") ? rel.newTarget.substring(3) : ("word/" + rel.newTarget);
              console.log(this.rels);
				
				if(rel.target != rel.newTarget)
				{
                  var filetypes = [".gif", ".png", ".jpeg", ".pdf"];
                  if (filetypes.some(x => rel.newTarget.includes(x))) {
                    zip.file(saveTo, rel.data);
                    relsXmlBuilder.push('<Relationship Id="' + rel.newId + '" Type="' + rel.type + '" Target="' + rel.newTarget + '"/>');
                  }
                    
                }
                /*else if (rel.filename.endsWith(".xml") && rel.filename === "styles.xml") {
                  var zipFile = zip.file(rel.zipPath);

                  String.prototype.replaceAll = function (search, replacement) {
                    var target = this;
                    return target.replace(new RegExp(search, 'g'), replacement);
                  };

                  var xml = this._utf8ArrayToString(rel.data).replaceAll(" />", "/>").substring(1);
                  xml = xml.substring(xml.indexOf("<"));
                  xml = xml.substring(xml.indexOf(">") + 1);

                  var closingTag = xml.substring(xml.lastIndexOf("</"));

                  var mergedXml = filesToSave[saveTo] || this._utf8ArrayToString(zipFile._data.getContent());
                  mergedXml = mergedXml.replace(closingTag, xml);
                  console.log(rel.filename, mergedXml);
                  filesToSave[saveTo] = mergedXml;

                } */
				else if(rel.filename.endsWith(".xml")) 
				{
                  var zipFile = zip.file(rel.zipPath);
                  String.prototype.replaceAll = function (search, replacement) {
                    var target = this;
                    return target.replace(new RegExp(search, 'g'), replacement);
                  };
					
					if((filesToSave[saveTo] || zipFile) && !rel.target.startsWith('theme/'))
					{
      //      var xml = this._utf8ArrayToString(rel.data).substring(1);
						//xml = xml.substring(xml.indexOf("<"));
      //                xml = xml.substring(xml.indexOf(">") + 1);

      //                if (rel.filename === "styles.xml" && xml.indexOf("<w:styles") > -1) {
      //                  xml = xml.substring(xml.indexOf(">") + 1);
      //                }

      //                if (rel.filename === "styles.xml") {
      //                  console.log(this._utf8ArrayToString(rel.data));
      //                }
						
						//var closingTag = xml.substring(xml.lastIndexOf("</"));
						
						//var mergedXml = filesToSave[saveTo] || this._utf8ArrayToString(zipFile._data.getContent());
      //                mergedXml = mergedXml.replace(closingTag, xml);
   
						//filesToSave[saveTo] = mergedXml;

                      var xml = convert.xml2json(this._utf8ArrayToString(rel.data), { compact: true, spaces: 0 });
                      var xmlOriginal = convert.xml2json(
                        filesToSave[saveTo] ||
                        this._utf8ArrayToString(zipFile._data.getContent()),
                        {
                          compact: true, spaces: 0
                        });

                      var mergedXml = Object.assign({}, JSON.parse(xml), JSON.parse(xmlOriginal));

                      var mergedRes = convert.json2xml(mergedXml, { compact: true, spaces: 0 });

                      filesToSave[saveTo] = mergedRes;
					}
					else
						filesToSave[saveTo] = this._utf8ArrayToString(rel.data);
				}
				else
					console.log("Cannot merge file " + rel.filename);
			}
			
			if(relsXmlBuilder.length > 0)
			{
				var relsXml = this._utf8ArrayToString(zip.file("word/_rels/document.xml.rels")._data.getContent());
				relsXmlBuilder.push('</Relationships>');
				relsXml = relsXml.replace('</Relationships>', relsXmlBuilder.join(''));
				zip.file("word/_rels/document.xml.rels", relsXml);
			}
			
			for(var path in filesToSave)
			{
				zip.file(path, filesToSave[path]);
			}
		}

      var addBody = this._utf8ArrayToString(zip.file("word/document.xml")._data.getContent());
      var watermark = `<w:sdtContent><w:r><w:pict w14:anchorId="15987AD4"><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas><v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/><v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/><v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/><v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/><v:f eqn="if @0 @4 21600"/><v:f eqn="mid @5 @6"/><v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/><v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @5"/></v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype><v:shape id="PowerPlusWaterMarkObject357831064" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:412.4pt;height:247.45pt;rotation:315;z-index:-251656704;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;calibri&quot;;font-size:1pt" string="LUONNOS"/><w10:wrap anchorx="margin" anchory="margin"/></v:shape></w:pict></w:r></w:sdtContent>`;

      var body = "";

      String.prototype.splice = function (idx, rem, str) {
        return this.slice(0, idx) + str + this.slice(idx + Math.abs(rem));
      };

      if (isDraft && zip.file("word/header3.xml") != null) {
        var addHeader = this._utf8ArrayToString(zip.file("word/header3.xml")._data.getContent());
         var newHeader = addHeader.splice(addHeader.indexOf(`</w:sdtPr>`), 0, watermark);
        zip.file("word/header3.xml", newHeader);
        body += `<w:footerReference w:type="default" r:id="rId15"/><w:headerReference w:type="first" r:id="rId16"/><w:footerReference w:type="first" r:id="rId17"/>`;
      }

      body += `<w:p w:rsidR="005F670F" w:rsidRDefault="005F79F5"><w:r><w:t>{@body}</w:t></w:r></w:p>`;

      var newBody = addBody.splice(addBody.indexOf(`</w:body>`), 0, body);

      zip.file("word/document.xml", newBody);

		var doc = new Docxtemplater().loadZip(zip);


		doc.setData({body: this._body.join(''), header: this._header.join(''), footer: this._footer.join('') });
		doc.render();
		
      var out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });
      return out;
    }

  var performance = global.performance || {}
  var performanceNow =
    performance.now ||
    performance.mozNow ||
    performance.msNow ||
    performance.oNow ||
    performance.webkitNow ||
    function () { return (new Date()).getTime() }

  this.hrtime = function(previousTimestamp) {
    var clocktime = performanceNow.call(performance) * 1e-3
    var seconds = Math.floor(clocktime)
    var nanoseconds = Math.floor((clocktime % 1) * 1e9)
    if (previousTimestamp) {
      seconds = seconds - previousTimestamp[0]
      nanoseconds = nanoseconds - previousTimestamp[1]
      if (nanoseconds < 0) {
        seconds--
        nanoseconds += 1e9
      }
    }
    return [seconds, nanoseconds]
  }
}
