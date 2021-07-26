# VBA regex world

This is the tool write in VBA, support for regex.

## About

This is the tool support for:
- Search with regex in list of file or folder. [in-progress]
- List up result to each cell.[in-progress]
- Replace by using regex, and output replacement information data sheet. [in-progress]
- Create an Excel function to direct use regex in Excel cells
- Filter text of files, by ignored or get by list of regex
- Clean up output information. [in-progress]

Why you use this tool:
- The UI of this tool is Excel. Easy to input
- The output is separate in each cell of Excel, easy to review and copy.
- Regex in VBA is using PCRE regex style, very common regex style, compatible with many other languages/library

Because VBA sourcecode are inside Excel file (binary file), then I need to use [vbaDeveloper](https://github.com/hilkoc/vbaDeveloper) to import, export and manage VBA source code version
Each time you open or close `vba-regex-world-dev.xlsm` the Excel tool, there will be an diaglog appears add for import and export VBA source.
If you only need to use the tool then you should use `vba-regex-world.xlsm`, this is`vba-regex-world-dev.xlsm` without [vbaDeveloper](https://github.com/hilkoc/vbaDeveloper)

## Usage



## VBA regex cheat sheet