# Introduction!

Are you looking for a way to use Word automation? Want to deal with Word documents programmatically? Go through this tip, it will help you to deal with Word Automation without Interop using C# and OpenXML.

After going through this tip, you can tell:

1. What is Open XML
2. Why to use Open XML
3. How to use OpenXML to create Word documents using C# and OpenXML API
4. Create Word table using OpenXML and C#

## Background
I have seen many developers struggling to deal with Word documents programmatically, There are couple of ways to deal with Word documents:

1. Using COM interop object (Winword instance)
2. Using OpenXML API (Do not have to install Word on machine)

## Using the Code

#### Things We Need

![Word](https://www.codeproject.com/KB/cs/994905/docx.jpg)
![openXML](https://www.codeproject.com/KB/cs/994905/OpenXML.jpg)

Before starting with the OpenXML cooking, we need the following things to be ready with us:

1. C# Visual Studio (2005+ version)
2. OpenXML API (can be downloaded from here [Open XML SDK 2.5 for Microsoft Office](http://www.microsoft.com/en-in/download/details.aspx?id=30425))
That's it. (*Wow!!! No word installation needed*)

#### Getting Started with OpenXML

Now a days, DOCX files are getting popular day by day, due to them being very light and faster in processing, DOCX is the magical result of ZIP and XML combination. So it is clear that if we able to manage XMLs, we will be able to manage DOCX too. For managing WordXML, we need some API and that API is known as Open XML SDK for Microsoft Office, MSDN Says "*API simplifies the task of manipulating Open XML packages and the underlying Open XML schema elements within a package. The Open XML SDK encapsulates many common tasks that developers perform on Open XML packages, so that you can perform complex operations with just a few lines of code.*"

#### Open XML Advantages over Interop
1 Open XML is an open standard for Word-processing documents, presentations, and spreadsheets that can be freely implemented by multiple applications on different platforms
2 The purpose of the Open XML standard is to de-couple documents created by Microsoft Office applications so that they can be manipulated by other applications independent of proprietary formats and without the loss of data.
3 As it is light weight, the processing is faster than interop objects
4 It has good Interoprability, Backwards Compatibility and Programmability
5 As it is Smaller File Size, it is to manage all variety of document stores, including Exchange servers, SharePoint, and of course network file storage.
6 It's a IS29500 standard, free for all to use, and extremely well documented

*Do you know You can unzip DOCX file*

Do you know you can unzip DOCX file? DOCX is the combination of several well structured .XML file, An Open XML file is stored in a ZIP archive for packaging and compression. You can view the structure of any Open XML file using a ZIP viewer, Open XML document is built of multiple document parts. The relationships between the parts are themselves stored in document parts, each typical DOCX file has the following different parts.

See the below image to know the different XML parts:

![Parts](https://www.codeproject.com/KB/cs/994905/TypicalDocument.jpg)

Body is the main part of the document and it has many different parts as shown in the above figure.

*Working with Paragraphs (First Assignment)*

Paragraphs is the most basic unit of block-level content within a WordprocessingML document, paragraphs are stored using the *<p>* element, Paragraph different sub elements like ParagraphProperties (Optional), Run and Text.

**Paragraph Properties**

Paragraph properties are used for the formatting of the text, some of the examples of paragraph properties are alignment, border, hyphenation override, indentation, line spacing, shading, text direction. The OXML SDK Paragraph properties class represents the *<pPr>* element

**Run**

The run element is provided to demarcate a region of text. The OXML SDK Run class represents the *<r>* element.

**Text**

This element contains actual Text of a document, With the <<span class="input">r</span>> element, the text (<<span class="input">t</span>>) element is the container for the text that makes up the document content.

*Start with the Code (Create new word document and write in it)*

Open Visual Studio and start with the first OpenXML assignment.

Create new Project/Application and add DLL reference (DLL should exist in Installed OpenXML API folder, e.g., C:\Program Files\Open XML SDK\V2.0\lib).

**DocumentFormat.OpenXml**

See the below snippet where we are creating new Word document with the help of OpenXML.

```javascript
using (WordprocessingDocument doc = WordprocessingDocument.Create
("D:\\test11.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
       {
           // Add a main document part.
           MainDocumentPart mainPart = doc.AddMainDocumentPart();

           // Create the document structure and add some text.
           mainPart.Document = new Document();
           Body body = mainPart.Document.AppendChild(new Body());
           Paragraph para = body.AppendChild(new Paragraph());
           Run run = para.AppendChild(new Run());

           // String msg contains the text, "Hello, Word!"
           run.AppendChild(new Text("New text in document"));
       }
```

In the above simple snippet:

- We have use 'WordProcessingDocument' class for creating new document
- Add MainDocumentPart in document
- Then append Body to main document part
- Then add Paragraph to Body element
- Then add Run to Paragraph element
- Then add Text to Run element

That's it. No need to save document anymore.

Now if you go and check for 'test11.docx', then you can see it contains text 'New text in document'.

Now try to unzip that Docx file, you will get below folder structure, you will get folders *_rels, docsProps, word* and *[Content_Types].xml* file.

Open Word folder and check document.xml. You will see the below snap:

![document](https://www.codeproject.com/KB/cs/994905/OpenXML_Structure.jpg)

In the above image, you can see *<w:body>* represents MainBody of the document, *<w:p>* is the paragraph element, *<w:r>* is the run element, *<w:t>* is the text element.

This is how OpenXML works.

### Facts

OpenXML is really an amazing thing, it fluently works with spreadsheets, charts, presentations, and Word processing documents. The Open XML file formats are useful for developers because they use an open standard and are based on well-known technologies: ZIP and XML.

*Special thanks*

Following are the referral links for OpenXML:

- [OpenXML1](https://msdn.microsoft.com/EN-US/library/office/bb456488.aspx)
- [OpenXML2](https://msdn.microsoft.com/EN-US/library/office/bb456487.aspx)
- [OpenXML3](http://blogs.msdn.com/b/ericwhite/archive/2008/10/20/eric-white-s-blog-s-table-of-contents.aspx)

*Finally*

OpenXML is not a single cup of tea, I am continuing with a different assignment on OpenXML in the next version of this article. Till then, enjoy this stuff. Suggestions and queries are always welcome.

Thanks

Prasad