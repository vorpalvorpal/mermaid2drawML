# print() shows the summary and returns the diagram invisibly

    Code
      print(d)
    Output
      <mermaid_word_diagram>: flowchart-v2
      Shapes: 10
      Shape ids: 5 to 15
      Orientation: portrait

# print() with custom margins

    Code
      print(mermaid_to_word(code, page_margins_in = c(top = 0.5, left = 0.75)))
    Output
      <mermaid_word_diagram>: flowchart-v2
      Shapes: 10
      Shape ids: 1 to 11
      Orientation: portrait
      Margins (in): top 0.5, bottom 1, left 0.75, right 1

