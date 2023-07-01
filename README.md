
# JiraWiki2Docx

JiraWiki2Docx is a Python 3 library that converts Jira wiki text to Microsoft Word (docx) format. It provides a convenient way to migrate Jira wiki content to Word documents, allowing users to preserve the formatting and structure of their Jira wiki pages.

## Installation

You can install jirawiki2docx using pip:

`pip install jirawiki2docx` 

## Usage

To use jirawiki2docx, follow these steps:

1.  Import the JiraWiki2Docx class:

	`from jirawiki2docx import JiraWiki2Docx` 

2.  Create an instance of the JiraWiki2Docx class, providing the Jira wiki text as input:
	```
	jira_text = "Your Jira wiki text here..."
	converter = JiraWiki2Docx(jira_text)
	``` 

3.  Parse the Jira wiki text and generate the Word document:

	`document = converter.parseJira2Docx()` 

4.  Optionally, save the resulting document to a file:

	`converter.parseJira2Docx(save_to_file=True, output_filename="output.docx")`

5. Alternatively, you can use an existing `docx.Document` object and append the converted Jira wiki text to it's content:

	```
	existing_document = docx.Document()
	jira_converter.parseJira2Docx(existing_document)
	```

## Example

Here's a complete example that demonstrates the usage of JiraWiki2Docx:

```
from docx import Document
from jirawiki2docx import JiraWiki2Docx

# Jira wiki text
jira_text = """
h1. Heading 1

This is a paragraph.

h2. Heading 2

* Item 1
* Item 2

h3. Heading 3

||Heading 1||Heading 2||
|Cell 1|Cell 2|

This is another paragraph.
"""

# Create a Document object
`document = Document()`

# Create an instance of JiraWiki2Docx and parse the Jira wiki text
converter = JiraWiki2Docx(jira_text, document)

# Parse Jira markup and convert to Word document with optional arguments to save created/modified document to file or just return as object
converter.parseJira2Docx(save_to_file=True, output_filename="output.docx")
```

The resulting Word document will contain the converted Jira wiki content with appropriate headings, lists, and tables.

## Supported Jira Wiki Syntax

JiraWiki2Docx supports the following Jira wiki syntax:

-   Headings (h1 to h6)
-   Unordered lists (*, -, #)
-   Tables (||)
-   Text effects (bold, italics, deleted, inserted, superscript, subscript, and color)

Please note that JiraWiki2Docx may not support all Jira wiki syntax and rendering options.

## License

JiraWiki2Docx is released under the MIT License. See the [LICENSE](https://github.com/deitar/jirawiki2docx/blob/main/LICENSE) file for more details.

## Conclusion

JiraWiki2Docx simplifies the process of converting Jira wiki markup to Word documents. Whether you need to share Jira content with stakeholders or generate printable reports, this library provides an easy-to-use solution. Give it a try and streamline your Jira content management and distribution process.