# PowerPoint Add-in: Send a PowerPoint document in chunks to a service

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#key-components-of-the-sample)
* [Description of the code](#description-of-the-code)
* [Build and debug](#build-and-debug)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions-and-comments)
* [Additional resources](#additional-resources)

## Summary
This sample shows how to use JavaScript in a PowerPoint task pane add-in to get the current presentation and slice it into chunks of data in user-defined sizes. The data could then be submitted to a service (such as an e-book publishing service).

## Prerequisites

This sample requires the following:  

  - PowerPoint for Windows 2013, 2016 or PowerPoint for Mac 2016  
  - Visual Studio 2015 (Update 3), with Microsoft Office Developer Tools.  
  - Any modern  browser such as Edge, Internet Explorer 11, Chrome or Safari.   

## Key components of the sample
The sample solution contains the following key files:

**EBookPublisher** project

- [EBookPublisher.xml](https://github.com/OfficeDev/PowerPoint-Add-in-JavaScript-SliceDataChunks/blob/master/EBookPublisher/EBookPublisherManifest/EBookPublisher.xml): The manifest file for the PowerPoint add-in.  
- [Adventure Works.ppt](https://github.com/OfficeDev/PowerPoint-Add-in-JavaScript-SliceDataChunks/blob/master/EBookPublisher/Adventure%20Works.pptx): Start Document with 1,024 slides. 
 
**EBookPublisherWeb** project

- [App/Home/Home.html](https://github.com/OfficeDev/PowerPoint-Add-in-JavaScript-SliceDataChunks/blob/master/EBookPublisherWeb/App/Home/Home.html). The HTML user interface that is displayed in the task pane. 
- [App/Home/Home.js](https://github.com/OfficeDev/PowerPoint-Add-in-JavaScript-SliceDataChunks/blob/master/EBookPublisherWeb/App/Home/Home.js). Logic that runs when the add-in is loaded. 


## Description of the code
The Adventure Works.pptx file is set as the **Start Document** property of the task pane add-in. The presentation is large enough (1,204 slides) to be sliced into a number of discrete chunks of data. 

The sample demonstrates:

- How to use JavaScript to retrieve the selected value from a drop-down list.
- How to use the `getFileAsync` method to slice the file into chunks of data of particular sizes.
- How to retrieve the data from each slice of the file by using the `getSliceAsync` method.


## Build and debug 

1. In Visual Studio, press F5 to build and deploy the sample add-in. The Adventure Works.pptx file opens in PowerPoint.
2. In the task pane add-in, choose a size for the data chunk.
3. Click the **Publish now!** button. 

The add-in displays the number of slices and the size of each slice, along with buttons you can use to view the content of each slice.

> Note: This sample displays the slice information to the user, but your add-in will probably send the data slices to a web service. The web service can then rebuild the presentation from the slices.

## Troubleshooting

- If the add-in starts with a blank presentation, ensure that the **Start Document** property of the EBookPublisher project is set to *Adventure Works.pptx* (not to *New PowerPoint Presentation*).
- If the presentation opens in read-only mode, click the **Enable editing** button.
- If the add-in does not appear in the task pane of the presentation, Choose **Insert > My Add-ins > EBook Publisher**.


## Questions and comments

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/PowerPoint-Add-in-JavaScript-SliceDataChunks/issues).
- Questions about Office Add-ins development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


## Additional resources

- [Office Add-in Documentation](http://dev.office.com/docs/add-ins/overview/office-add-ins)
- [Get the whole document from an add-in for PowerPoint or Word](https://dev.office.com/docs/add-ins/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word)
- [Document.getFileAsync method](https://dev.office.com/reference/add-ins/shared/document.getfileasync)
- [File.getSliceAsync method](https://dev.office.com/reference/add-ins/shared/document.getfileasync)
- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.
