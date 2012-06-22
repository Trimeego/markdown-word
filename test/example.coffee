md = require("../lib/markdown-word")
md.documentFromFile "example.md", "./example.docx", ((err, data) ->
  console.log err, JSON.stringify(data, null, 2)
), 0

