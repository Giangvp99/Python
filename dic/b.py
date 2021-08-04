test = "tester.rtf"
out_file = open(test, "w")
out_file.write(
    """{\\rtf1
This is \\b Bold  \\b0\line\
}"""
)
out_file.close()
