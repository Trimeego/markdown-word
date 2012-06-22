markdown = require( "markdown" ).markdown
fs = require("fs-extra")
path = require("path")
temp = require("temp")
zip = require("./zip/janzip.js")
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


millisOld = 0
counter = 0
id = 1

resolvePath = (from, to) ->
  out = 
    type: "absolute"
    path: to

  if from
    if from.href
      # url
      toUrl = url.parse(to)
      if not toUrl.host
        #relative
        out = 
          type: "relative"
          path: url(from, to)
    else
      # file path
      if to.indexOf(".") is 0
        #relative path
        out = 
          type: "relative"
          path: path.resolve(from,to)      
  out


createUid = () ->
  protectRollover = false
  millis = new Date().getTime() - 1262304000000 * Math.pow(2, 12)
  id2 = id * Math.pow(2, 8)
  uid = Math.abs(millis + id2 + counter + Math.round(Math.random(10)* 100))

addRelationships = (rels) ->
  for rel in rels
    if rel
      relAttribs = 
        "Id": rel.relId
        "Type": rel.type
        "Target": rel.href

      if rel.targetMode
        relAttribs["TargetMode"] = rel.targetMode


      relationships["Relationships"].push
        "Relationship": [
          _attr: relAttribs
        ]
  relationships 

processChildElements = (list, children, levelOffset, rels, rootPath)->
  for el in list
    objectResults = processMarkdownObject(el, levelOffset)
    if Object.prototype.toString.call( objectResults ) is '[object Array]'
      for r in objectResults
        if Object.prototype.toString.call( r.markup ) is '[object Array]'
          Array::push.apply children, r.markup 
        else
          children.push r.markup 
        Array::push.apply rels, r.relationships
    else
      if Object.prototype.toString.call( objectResults.markup ) is '[object Array]'
        Array::push.apply children, objectResults.markup 
      else
        children.push objectResults.markup 
      
      Array::push.apply rels, objectResults.relationships



processMarkdownObject = (object, levelOffset, rootPath) ->
  children = []
  rels = [] 
  if Object.prototype.toString.call( object ) is '[object Array]'
    [objectType, elements...] = object
  else
    objectType = "text"
    elements = object

  switch objectType
    when "markdown"
      fragments = []
      processChildElements elements, fragments, levelOffset, rels, rootPath
      # children.push documentSectionProps
      # out = 
      #   "w:document": [
      #     _attr: documentAttrs
      #   ,
      #     "w:body": children
      #   ]
      out = fragments

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

      processChildElements childElements, children, levelOffset, rels, rootPath

      out = 
        "w:p": children

    when "para"
      processChildElements elements, children, levelOffset, rels, rootPath
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
        processChildElements itemElements, children, levelOffset, rels, rootPath

        out.push     
          "w:p": children

    when "numberlist"
      rid = createUid()
      out = []
      count = 0
      for el in elements
        count++
        children = [
          "w:pPr": [
            "w:pStyle": [
              _attr:
                "w:val": "ListParagraph"

            ]
          ,
            "w:numPr": [
              "w:ilvl":
                _attr:
                  "w:val": "0"
            ,
              "w:numId":
                _attr:
                  "w:val": "#{rid}"
            ]
          ,
            "w:ind":
              _attr:
                "w:left": "720"
                "w:hanging": "360"
          ]
        ]
        [itemType, itemElements...] = el
        processChildElements itemElements, children, levelOffset, rels, rootPath
        out.push     
          "w:p": children

    when "strong"
      children.push 
        "w:rPr": [
          "w:b": ""
        ]
      processChildElements elements, children, levelOffset, rels, rootPath
      children.push
        "w:r": [
          "w:t": [
            _attr: 
              "xml:space":"preserve"  
          ,          
            " "
          ]
        ] 

      out = 
        "w:r": children

    when "em"
      children.push 
        "w:rPr": [
          "w:i": ""
        ]
      processChildElements elements, children, levelOffset, rels, rootPath
      children.push
        "w:r": [
          "w:t": [
            _attr: 
              "xml:space":"preserve"  
          ,          
            " "
          ]
        ] 

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

    when "code_block"
      out = [
        "w:p": ""
      ]
      lines = elements[0].split "\n"
      for l in lines
        # r = processMarkdownObject ["code_line", l], levelOffset, rootPath  
        # rels.push r.relationships
        out.push 
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
                l
              ]
            ]
          ]    

      
      out.push 
        "w:p": "" 

    when "link"
      rp = resolvePath rootPath, elements[0].href
      href = rp.path
      rid = "rId" + createUid()
      title = elements[1]
      rels.push 
        href: href
        rid: rid
        title: title
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" 
        targetMode: "External"
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
      rid = "rId" + createUid()
      rels.push 
        href: elements[0].href
        rid: rid
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 

      out = [
        "w:pict": [
          "v:shape": [
            "v:imagedata": [
              _attr:
                "r:id": rid
            ]
          ]
        ]  
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

  r = 
    markup: out
    relationships: rels

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


buildDocumentObjectFromFragments = (fragments) ->
  docObjFragments = []
  docRelationships = []

  if fragments.markup
    Array::push.apply docObjFragments, fragments.markup
    if fragments.relationships
      Array::push.apply docRelationships, fragments.relationships
  else
    for fragment in fragments
      Array::push.apply docObjFragments, fragment.markup
      Array::push.apply docRelationships, fragment.relationships
    # docObjFragments.push documentSectionProps
    
  for r in docRelationships
    relationships["Relationships"].push 
      "Relationship": [
        _attr: 
          "Id": r.rid
          "Type": r.type
          "Target": r.href
          "TargetMode": r.targetMode
      ]
  
  out = 
    relationships: relationships
    markup:
      "w:document": [
        _attr: documentAttrs
      ,
        "w:body": docObjFragments
      ]
  out

buildDocumentFromFragments = (fragments, outputFile) =>
  targetFiles = []
  document = buildDocumentObjectFromFragments(fragments)
  
  documentXML = 
      """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      #{XML(document.markup)}
      """    

  relationshipsXML = 
      """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      #{XML(addRelationships(document.relationships))}
      """    


  imageCount = 0
  for image in document.relationships when image.type is "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    imageCount++
    targetFiles.push
      name: "word/media/#{imageCount}"
      path: image.href
      compression: 'store'


  walk templatePath, (err, files) ->

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

buildDocumentFromMarkdown = (inputMarkdown, outputFile, levelOffset, rootPath) =>
  if not outputFile
    info = temp.openSync "md2word"
    outputFile = info.path

  markdownJSON = markdown.parse( inputMarkdown )
  fragments = processMarkdownObject markdownJSON, levelOffset, rootPath
  buildDocumentFromFragments fragments, outputFile, levelOffset, rootPath



markdownFromFile = (filepath, callback) =>
  fs.readFile filepath, (err, data) ->
    callback(err,data)


markdownFromUrl = (fileUrl, callback) =>
    inputMarkdown = ""

    if fileUrl.split(":")[0] is "https"
      port = 443
      lib = https
    else
      port = 80
      lib = http

    options = 
      host: url.parse(fileUrl).host
      port: port
      path: url.parse(fileUrl).pathname

    lib.get options, (res) ->
      res.on 'data', (data) ->
        inputMarkdown += data
      res.on 'end', () -> 
        callback(null, inputMarkdown)


module.exports = 
  documentFromFile: (filepath, outputFile, callback, levelOffset=0) ->
    rootPath = path.dirname(path.resolve(filepath))
    markdownFromFile filepath, (err, data) ->
      out = buildDocumentFromMarkdown data, outputFile, levelOffset, rootPath
      callback(null, out)

  documentFromMarkdown: (inputMarkdown, outputFile, callback, levelOffset=0, rootPath=null) ->
    out = buildDocumentFromMarkdown inputMarkdown, outputFile, levelOffset, rootPath
    callback(null, out)

  documentFromUrl: (fileUrl, outputFile, callback, levelOffset=0) ->
    rootPath = url.parse(fileUrl)
    markdownFromUrl fileUrl, (err, data) ->
      out = buildDocumentFromMarkdown data, outputFile, levelOffset, rootPath;
      callback(null, out);

  documentFromFragments: (fragments, outputFile, callback) ->
      out = buildDocumentFromFragments fragments, outputFile
      callback(null, out);

  fragmentsFromFile: (filepath, callback, levelOffset=0) ->
    rootPath = path.dirname(path.resolve(filepath))    
    markdownFromFile filepath, (err, data) ->
      markdownJSON = markdown.parse( data )

      fragments = processMarkdownObject markdownJSON, levelOffset, rootPath
      callback(null, fragments)

  fragmentsFromMarkdown: (inputMarkdown, callback, levelOffset=0, rootPath=null) ->
    markdownJSON = markdown.parse( inputMarkdown )
    fragments = processMarkdownObject markdownJSON, levelOffset, rootPath
    callback(null, fragments)

  fragmentsFromUrl: (fileUrl, callback, levelOffset, rootPath) ->
    rootPath = url.parse(fileUrl)    
    markdownFromUrl fileUrl, (err, data) ->
      markdownJSON = markdown.parse( data )
      fragments = processMarkdownObject markdownJSON, levelOffset, rootPath
      callback(null, fragments)
