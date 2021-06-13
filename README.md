# Azure function Word DOCX Generator

POSTing the JSON to the function will download the DOCX template from the _TEMPLATE_URL parameter, and attempt to replace any tokens in the format of ((token)) with the values provided in the JSON.

## Languages

Two implementations of the same code, as C# for traditional deployment, or CSX script.

### C#

Project contains C# source targetting V3 Azure Functions.

### CSX script based

* Replace run.csx
* Upload function.proj

## How to run it

HTTP POST to the function

```
{
    "_TEMPLATE_URL" : "https://storageaccountrgpura359.blob.core.windows.net/tech-challenge-templates/Technical Assignment Template.docx",
    "username" : "xyz@dev",
    "password" : "12345",
    "uri" : "https://somewhere.com",
}

## Content returned

The return value is file of application/vnd.openxmlformats-officedocument.wordprocessingml.document.