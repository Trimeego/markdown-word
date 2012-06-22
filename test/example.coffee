md = require("../lib/markdown-word")
md.documentFromFile __dirname + "/example.md", __dirname + "/example.docx", ((err, data) ->
  console.log err, JSON.stringify(data, null, 2)
), 0

