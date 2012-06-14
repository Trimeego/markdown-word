markdown = require( "markdown" ).markdown
fs = require("fs-extra")
temp = require("temp")
zip = require("node-native-zip")
XML = require("xml")
http = require("http")
https = require("https")
url = require("url")


documentAttrs =  
  "xmlns:wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" 
  "xmlns:mo": "http://schemas.microsoft.com/office/mac/office/2008/main" 
  "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006" 
  "xmlns:mv": "urn:schemas-microsoft-com:mac:vml" 
  "xmlns:o": "urn:schemas-microsoft-com:office:office" 
  "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
  "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math" 
  "xmlns:v": "urn:schemas-microsoft-com:vml" 
  "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" 
  "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" 
  "xmlns:w10": "urn:schemas-microsoft-com:office:word" 
  "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" 
  "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml" 
  "xmlns:wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" 
  "xmlns:wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk" 
  "xmlns:wne": "http://schemas.microsoft.com/office/word/2006/wordml" 
  "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" 
  "mc:Ignorable": "w14 wp14"

documentSectionProps = 
  "w:sectPr": [
    "w:pgSz":
      _attr:
        "w:w": "12240"
        "w:h": "15840"            
  ,  
    "w:pgMar":
      _attr:
        "w:top": "1440"
        "w:right": "1800"
        "w:bottom": "1440"
        "w:left": "1800"
        "w:header": "720"
        "w:footer": "720"
        "w:gutter": "0"
  ,
    "w:cols":
      _attr:
        "w:space": "720"            
  ,
    "w:docGrid":
      _attr:
        "w:linePitch": "360"          
  ]


relationships = 
  "Relationships": [
    _attr: {
      "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
    }
  ,
    "Relationship": [
      _attr: {
        "Id": "rId1"
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" 
        "Target": "numbering.xml"        
      }
    ]  
  ,
    "Relationship": [
      _attr: {
        "Id": "rId2" 
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        "Target": "styles.xml"
      }
    ]
  ,
    "Relationship": [
      _attr: {
        "Id": "rId3" 
        "Type": "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
        "Target": "stylesWithEffects.xml"
      }
    ]
  ,
    "Relationship": [
      _attr: {
        "Id": "rId4" 
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
        "Target": "settings.xml"
      }
    ]
  ,
    "Relationship": [
      _attr: {
        "Id": "rId5" 
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
        "Target": "webSettings.xml"
      }
    ]
  ,
    "Relationship": [
      _attr: {
        "Id": "rId6" 
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
        "Target": "fontTable.xml"
      }
    ]
  ,
    "Relationship": [
      _attr: {
        "Id": "rId7" 
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        "Target": "theme/theme1.xml"
      }
    ]    
  ]



addRelationship = (link) ->
  relId = "rId#{relationships["Relationships"].length+1}"
  relationships["Relationships"].push
    "Relationship": [
      _attr: {
        "Id": relId
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" 
        "Target": link
        "TargetMode": "External"
      }
    ]  
  relId

processChildElements = (list, children, levelOffset)->
  for el in list
    objectResults = processMarkdownObject(el, levelOffset)
    if Object.prototype.toString.call( objectResults ) is '[object Array]'
      for r in objectResults
        children.push r
    else
      children.push objectResults 


processMarkdownObject = (object, levelOffset) ->
  children = [] 
  if Object.prototype.toString.call( object ) is '[object Array]'
    [objectType, elements...] = object
  else
    objectType = "text"
    elements = object

  switch objectType
    when "markdown"
      processChildElements elements, children, levelOffset
      children.push documentSectionProps
      out = 
        "w:document": [
          _attr: documentAttrs
        ,
          "w:body": children
        ]

    when "header"
      [attributes, childElements...] = elements
      level = Number(attributes.level)+levelOffset
      children.push 
        "w:pPr": [
          "w:pStyle": [
            _attr:
              "w:val": "Heading#{level}"
          ]
        ]

      processChildElements childElements, children

      out = 
        "w:p": children

    when "para"
      processChildElements elements, children
      out = 
        "w:p": children
    
    when "bulletlist"
      out = []
      count = 0
      for el in elements
        count++
        children = [
          "w:pPr": [
            "w:pStyle":
              _attr:
                "w:val": "ListParagraph"
          ,
            "w:numPr": [
              "w:ilvl":
                _attr:
                  "w:val": "0"
            ,
              "w:numId":
                _attr:
                  "w:val": "2"
            ]
          ]
        ]
        [itemType, itemElements...] = el
        processChildElements itemElements, children
        out.push     
          "w:p": children

    when "numberlist"
      out = []
      count = 0
      for el in elements
        count++
        children = [
          "w:pPr": [
            "w:pStyle":
              _attr:
                "w:val": "ListParagraph"
          ,
            "w:numPr": [
              "w:ilvl":
                _attr:
                  "w:val": "0"
            ,
              "w:numId":
                _attr:
                  "w:val": "1"
            ]
          ]
        ]
        [itemType, itemElements...] = el
        processChildElements itemElements, children
        out.push     
          "w:p": children

    when "strong"
      children.push 
        "w:rPr": [
          "w:b": ""
        ]
      processChildElements elements, children
      out = 
        "w:r": children

    when "inlinecode"
      out =     
        "w:r": [
          "w:t": [
            _attr: 
              "xml:space":"preserve"  
          ,          
            " "
          ]
        ,          
          "w:rPr": [
            "w:rFonts": [
              _attr: 
                "w:ascii": "Courier New" 
                "w:hAnsi": "Courier New" 
                "w:cs": "Courier New"
            ]
          ,
            "w:sz": [
              _attr: 
                "w:val": "20"
            ]
          ,
            "w:szCs": [
              _attr: 
                "w:val": "20"
            ]
          ]
        ,
          "w:t": [
            _attr: 
              "xml:space":"preserve"  
          ,          
            elements[0]
          ]
        ]

    when "code_line"
      out =     
        "w:p": [
          "w:pPr": [
            "w:pStyle": [
              _attr:
                "w:val": "CodeBlock"
            ]
          ]
        ,
          "w:r": [
            "w:t": [
              _attr: 
                "xml:space":"preserve"  
            ,          
              elements[0]
            ]
          ]
        ]    

    when "code_block"
      out = [
        "w:p": ""
      ]
      lines = elements[0].split "\n"
      for l in lines
        out.push processMarkdownObject ["code_line", l], levelOffset     
      
      out.push 
        "w:p": "" 

    when "link"
      href = elements[0].href
      rid = addRelationship(href)
      title = elements[1]
      out = [
        "w:r": [
          "w:t": [
            _attr: 
              "xml:space":"preserve"  
          ,          
            " "
          ]
        ]
      , 
        "w:hyperlink": [
          _attr: 
            "r:id": rid 
            "w:history": "1"
        ,
          "w:r": [
            "w:rPr": [
              "w:rStyle": [
                _attr:
                  "w:val": "Hyperlink"
              ]
            ,
              "w:t": [
                title
              ]
            ]
          ]
        ]
      ]

    when "img"
      out = 
        "w:r": [
          "w:t": "<<< #{objectType} not yet implemented>>>"
        ]

    when "text"
      #if it get's here it should just be a text element
      out = 
        "w:r": [
          "w:t": elements
        ]

    else
      out = 
        "w:r":[
          "w:t": "<<< #{objectType} not yet implemented>>>"
        ]

  out


walk = (dir, done) ->
  results = []
  fs.readdir dir, (err, list) ->
    return done(err)  if err
    pending = list.length
    return done(null, results)  unless pending
    list.forEach (file) ->
      file = dir + "/" + file
      fs.stat file, (err, stat) ->
        if stat and stat.isDirectory()
          walk file, (err, res) ->
            results = results.concat(res)
            done null, results  unless --pending
        else
          results.push file
          done null, results  unless --pending


# prepare the temp directory

templatePath = __dirname + "/template"



buildDocument = (inputMarkdown, outputFile, levelOffset) =>
  if not outputFile
    info = temp.openSync "md2word"
    outputFile = info.path

  markdownJSON = markdown.parse( inputMarkdown )
  document = processMarkdownObject markdownJSON, levelOffset
  documentXML = 
      """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      #{XML(document)}
      """    

  relationshipsXML = 
      """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      #{XML(relationships)}
      """    


  walk templatePath, (err, files) ->

    targetFiles = []
    for file in files
      if file.indexOf(".DS_Store") is -1 
        targetFiles.push
          name:  file.replace("#{templatePath}/", "")
          path:  file
          compression: 'store'


    archive = new zip()

    temp.open "md2word", (err, info) ->
      fs.write info.fd, new Buffer(documentXML, "utf8").toString()
      fs.close info.fd, (err) ->
        targetFiles.push
          name:  "word/document.xml"
          path:  info.path
          compression: 'store'

        temp.open "md2word", (err, info) ->
          fs.write info.fd, new Buffer(relationshipsXML, "utf8").toString()
          fs.close info.fd, (err) ->
            targetFiles.push
              name:  "word/_rels/document.xml.rels"
              path:  info.path
              compression: 'store'

            archive.addFiles targetFiles, (err) ->

              if err 
                console.log "err while adding files", err
              else
                buff = archive.toBuffer (result) ->
                  fs.writeFile outputFile, result, () ->

  outputFile

module.exports = 
  fromFile: (filepath, outputFile, callback, levelOffset=0) ->
    fs.readFile filepath, (err, data) ->
      throw err if err 
      out = buildDocument data, outputFile, levelOffset
      callback(null, out)

  fromMarkdown: (inputMarkdown, outputFile, callback, levelOffset=0) ->
    out = buildDocument inputMarkdown, outputFile, levelOffset
    callback(null, out)

  fromUrl: (fileUrl, outputFile, callback, levelOffset=0) ->
    inputMarkdown = ""

    if fileUrl.split(":")[0] is "https"
      port = 443
      lib = https
    else
      port = 80
      lib = http

    options = 
      host: url.parse(fileUrl).host,
      port: port,
      path: url.parse(fileUrl).pathname

    lib.get options, (res) ->
      res.on 'data', (data) ->
        inputMarkdown += data
      res.on 'end', () ->
        out = buildDocument inputMarkdown, outputFile, levelOffset
        callback(null, out)

