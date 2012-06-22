md = require("../lib/markdown-word")

input = """
## BaseView

Provides some basic rendering options that other views can inherit

### Options

- **templateHTML : string**  The template HTML.  Should be overridden
- **className : string**  The class to render the view into.  Should be overridden
- **layout : string**  Either a section based or row based layout.  Should be overridden

### Installation

1.  Open `Browser`
2.  Go To **WEB Link** [Test](http://www.google.com)
3.  Web Site
4.  Look at eh 

### Bulleted Installation

* Open `Browser`
* Go To **WEB Link** [Test](http://www.google.com)
* Web Site
* Look at eh 


### Usage

#### Using a row based layout

Rows based layouts are useful for simple views.  In this case, the layout should have a collection of sections at the root

    class NewView extends BaseView
      model: new Backbone.Model()
      layout: 
        rows: [ 
          fields: [
            label: "test"
            dataField: "test"
            css: "eight phone-four columns"
            type: "text"
            required: false
            readOnly: false
          ]
        ]



#### Using a section based layout

Section based layouts are useful for providing sectioned views.  In this case, the layout should have a collection of sections at the root

    class NewView extends BaseView
      model: new Backbone.Model()
      layout: 
        sections:[
          title: "Test"
          rows: [ 
            fields: [
              label: "test"
              dataField: "test"
              css: "eight phone-four columns"
              type: "text"
              required: false
              readOnly: false
            ]
          ]
        ]

"""


md.documentFromMarkdown input, __dirname + "/document-from-markdown.docx", (err, file) ->
  console.log file  

# md.documentFromFile __dirname + "/example.md", __dirname + "/example.docx", (err, file) ->
#   console.log file  

# md.documentFromUrl "https://raw.github.com/theironcook/Backbone.ModelBinder/master/README.md", __dirname + "/document-from-url.docx", (err, file) ->
#   console.log file  

