# ChatGPT Excel Functions

A set of functions for communicating with the OpenAI ChatGPT API in Microsoft Excel and perform various taasks. This was tested under the latest version of Microsoft Excel. Adjustments may be needed to run with older versions.

# Update 3/10/2023

Added ChatGPTQuerySelection() function.  Allows you to ask query ChatGPT about the data in a selection of cells.

![ChatGPTQuerySelection Video](https://img.youtube.com/vi/VrmhIFKKYBA/0.jpg)
(https://www.youtube.com/watch?v=VrmhIFKKYBA)

# Update (3/9/2023)

Added a ChatGPTList() function.  Pass in something you want a list of like in the video below, you can have the list output vertically and horizontally.

ChatGPTList demo:

![ChatGPTList Video](https://img.youtube.com/vi/bLxEHdxFb9k/0.jpg)
(https://www.youtube.com/watch?v=bLxEHdxFb9k)

# Update (3/8/2023)

Added a ChatGPTQuickFill() function, you pass in the number of cells you want to look backward for on the fill (if the cells don't exist or are empty they are ignored).  No prompt engineering required, the prompt is engineered automatically.  

Here's a demo video of the new Quick Fill:

![ChatGPT Quick Fill Video](https://img.youtube.com/vi/t9rUA67DV0E/0.jpg)
(https://www.youtube.com/watch?v=t9rUA67DV0E)

## Demo

![ChatGPT In Excel Demo](demo.gif)

## Installation

These instructions may only apply to recent versions of Microsoft Windows. You will need:

1. Open Microsoft Excel and click on the "Developer" tab.
2. Click on the "Visual Basic" button.
3. In the "Microsoft Visual Basic for Applications" window, click on "Insert" and then "Module".
4. Replace the existing code with the code in the "chatgpt-excel-function.vba" file.
5. Replace `sk-YOUR-CHATGPT-KEY-HERE` with your own OpenAI API key.

## Usage

Prompt chatGPT in a cell by typing `=ChatGPT(prompt)`, where `prompt` is the text you want to send to the ChatGPT API.

Quick fill cells using ChatGPT without a prompt using `=ChatGPTQuickFill(optional titleCell, optional contextCell)` where title row and context row default to row 1, and column a.  A prompt will be automatically generated.

Insert lists with ChatGPT using `=ChatGPTList(topic, optional horizontal)` The list will be created based on the passed in topic, by default the list will be placed vertically.  

Query cells with ChatGPT Using `=ChatGPTQuerySelection(query, selection, optional isheader)` This allows you query a selection of cells, you can set isHeader to true and it will use the first row as topic headers.  Currently this uses JSON, but that may change or have another option because of higher token count.  ChatGPT though seems to work best with JSON data rather than other formats tested. 

## Troubleshooting

If you receive an error related to `WinHttp.WinHttpRequest.5.1`, it means that the `WinHttp` library is not registered on your computer. To resolve this issue, you can try re-registering the library by running the following command in an elevated command prompt:

regsvr32 %systemroot%\system32\winhttp.dll

If the error persists, you may need to reinstall the library. You can download it from the Microsoft website.

## Note
The `UnescapeString` function is used to format the response from the API into a human-readable format. If you encounter any errors or unexpected output, please refer to the comments in the code for further assistance.

## Disclaimer
The authors and contributors of this program provide it as-is, without any warranties or guarantees. They cannot be held responsible for any damages resulting from the use of this program.

## License
This program is licensed under the MIT license.

## Author
Johann Dowa
