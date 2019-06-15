# Convert DOC to JSON with bible verse references

### Task
You're asked to write a nodejs program that takes a Word doc file name as input, it parses the text and output a .json file.

Please use the 3 sample doc files under "docs" folder for testing.
e.g.
```
node Parse.js 000Test.doc
````
Output file is 
```
000Test.json
```

### Json file format
It's a nested Json name/value format.
```
{
	"filename": "000Test.doc",
	"content": [{
		"name": "title",
		"value": "Romans: Good News from God"
	}, {
		"name": "title",
		"value": "An Overview"
	}, {
		"name": "paragraph",
		"value": [{
			"name": "bold_underline_text",
			"value": "THE THEME"
		}, {
			"name": "bold_text",
			"value": ": "
		}, {
			"name": "text",
			"value": "God graciously and sovereignly offers the full benefits of the gospel to those who believe in Christ and become established in Him (1:16–17; 4:3–5; 1:11; 16:25).",
			"bible": [{
				"verse": "1:16-17"
			}, {
				"verse": "4:3–5"
			}, {
				"verse": "1:11"
			}, {
				"verse": "16:25"
			}]
		}]
	}, {
		"name": "paragraph",
		"value": [{
			"name": "bold_underline_text",
			"value": "KEY WORDS"
		}, {
			"name": "bold_text",
			"value": ": "
		}, {
			"name": "text",
			"value": "The key word is “gospel/preach the gospel” ("
		}, {
			"name": "italic_text",
			"value": "euangelion"
		}, {
			"name": "text",
			"value": ", twelve times)"
		}]
	}, {
		"name": "paragraph",
		"value": [{
			"name": "text",
			"value": "Set apart (Acts 13:2) likely the work of the HS making 1:1 Trinitarian",
			"bible": [{
				"book": "Acts",
				"verse": " 13:2"
			}]
		}]
	}]
}
```