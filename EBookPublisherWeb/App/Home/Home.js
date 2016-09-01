/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />

(function () {
    "use strict";

    // This function is run when the add-in is ready to start interacting with the host application
    // It ensures the DOM is ready before adding a click handlers to the sendData button
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#sendData').click(function () { transmitChunk(); });

            $("#showChunkDialog").dialog({
                modal: true,
                open: function () {
                    $(document.body).css({ 'cursor': 'default' });
                },
                autoOpen: false
            });
        });
    };

    function transmitChunk() {
        // Get a reference to the <DIV> where we will write the outcome of our operation
        var report = document.getElementById('transmissionReport');

        // Get the selected value in the drop-down. We will use this value as a parameter
        // in the getFileAsync method. NOTE: The display values in the drop-down are shown in MB,
        // but the actual value returned is in Bytes (see the HTML for the drop-down in PowerPointEBookPublisher.html)
        var chunksize = document.getElementById('chunkSize').value;

        // Initialize the variable that we will use to store the number of slices returned
        // by the getFileAsync method
        var totalSlices = 0;

        // Remove all nodes from the transmissionReport <DIV> so we have a clean space to write to
        while (report.hasChildNodes()) {
            report.removeChild(report.lastChild);
        }

        // Now we can begin the process.
        // Step 1 is to call the getFileAsync method. Because this is PowerPoint,
        // the first fileType parameter must be the string value "compressed" or the enumerated equivalent of
        // Office.FileType.Compressed. Note: If this was Word, you could choose between "compressed" or "text".
        // The second parameter is the size of chunks that we want to slice the document into.
        // The value is in Bytes, returned from the chunkSize drop-down list.
        // The reason we use parseInt is that we need to provide an integer value, and the drop-down values are strings.
        // When the method returns, the function that is provided as the third parameter will run.
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: parseInt(chunksize) }, function (result) {
            if (result.status == "succeeded") {
                // If the getFileAsync call succeeded, then
                // result.value will return a valid File object, which we'll
                // hold in the currentFile variable.
                var currentFile = result.value;

                // Now we can start accessing the properties of the returned File object.
                // First, we'll create a <DIV> and tell the user how big the file is (in MB). The size property is actually
                // returned in Bytes, so we first need to convert that to MB, and then use that value in our own function
                // called trimSize which simply ensures the value is returned to two decimal places. Otherwise we could be displaying
                // values such as 3.83872637 (which is not very tidy :-)
                var fileData = document.createElement("div");
                var fileDataText = document.createTextNode("Total file size: " + trimSize(parseFloat((currentFile.size / 1024) / 1024)) + " MB");
                fileData.appendChild(fileDataText);
                report.appendChild(fileData);

                // Then we'll use the sliceCount property of the fileObject to tell the user how many slices there are
                totalSlices = currentFile.sliceCount;
                var sliceData = document.createElement("div");
                var sliceDataText = document.createTextNode("Number of slices: " + totalSlices);
                sliceData.appendChild(sliceDataText);
                report.appendChild(sliceData);

                // Now we'll actually do something with each slide
                for (var slice = 0; slice < totalSlices; slice++) {
                    // We'll call the getSliceAsync method of the File object, and pass in the
                    // integer in the above 'for' loop as the first parameter. This is simply an index
                    // which indicates which slice to get
                    var currentSlice = currentFile.getSliceAsync(slice, function (result) {
                        if (result.status == "succeeded") {
                            // If the getSliceAsync call succeeded, then
                            // result.value will return a valid Slice object, from which we'll
                            // access various properties.
                            // The first thing we'll do is get the actual slice data. This is effectively what can be
                            // used to rebuild a file, slice-by-slice. In our case, we'll encode it as Base64
                            // and store it temporarily in the following variable
                            // If using Internet Explorer, requires Internet Explorer 10.
                            var encData = btoa(result.value.data);

                            // The next thing we'll do is get the slice size and report it to the user.
                            // In this case, we're retrieving the 'size' in Bytes, so we first need to convert that to KB,
                            // and then use that value in our own function called trimSize which simply ensures the value is
                            // returned to two decimal places. Otherwise we could be displaying
                            // values such as 243.83872637 (which is not very tidy :-)
                            // We're also retrieving the index value of the slice, so that we can display the following pattern
                            // to the user:
                            // "Sending slice 1: 256.00 KB"
                            // "Sending slice 2: 256.00 KB"
                            // "Sending slice 3: 237.22 KB"
                            var sizeData = document.createElement("div");
                            var sizeDataDetails = document.createTextNode("Sending slice " + (result.value.index + 1) + ": " + trimSize(result.value.size / 1024) + " KB");
                            sizeData.appendChild(sizeDataDetails);
                            sizeData.appendChild(document.createElement("br"));

                            // Now for some fun: We'll take the actual raw data of each slice and let the user
                            // actually see it!!
                            // We'll create a button and wire up its onclick attribute on-the-fly
                            // so that it passes in the raw data of the slice to our function called showChunk.
                            // NOTE: The showChunk function is near the end of this script file
                            var rawData = document.createElement("button");
                            var rawDataDetails = document.createTextNode("View raw data");
                            rawData.setAttribute("class", "ms-Button");
                            rawData.setAttribute("onclick", "showChunk('" + encData + "');");
                            var label = document.createElement("span");
                            label.setAttribute("class", "ms-Button-label");
                            label.appendChild(rawDataDetails);
                            rawData.appendChild(label);
                            sizeData.appendChild(rawData);
                            report.appendChild(sizeData);

                            // Finally, we'll tell the user when we're finished.
                            // Basically, the slice variabe in the 'for' loop will have
                            // been incremented to one more than the highest indexed slice object returned
                            // from the getSliceAsync method on the last go through the loop.
                            if ((result.value.index + 1) == slice) {
                                var endMessage = document.createElement("div");
                                var endText = document.createTextNode("File has been sent!");
                                endMessage.appendChild(endText);
                                report.appendChild(endMessage);
                            }
                        }
                        else {
                            // This runs if the getSliceAsync method does not return a success flag
                            app.showNotification("Error", result.error.message);
                        }
                    });
                }
                // We're done with the File object, so we'll release its handle and thereby free up the
                // memory we've been using to slice it
                currentFile.closeAsync();
            }
            else
                // This runs if the getFileAsync method does not return a success flag
                app.showNotification("Error", result.error.message);
        });
    }
})();

// This function handles the Click events of the buttons we've added in the transmitChunk function above.
// It shows the raw data for the appropriate Slice object in a jQuery dialog
function showChunk(dataChunk) {
    $(document.body).css({ 'cursor': 'wait' });
    $("#showChunkDialog").html(dataChunk);
    $("#showChunkDialog").dialog("open");
}

// Very simple function for taking a string that looks like a number with potentially many decimal places
// and returns a string that looks like a number with only two decimal places.
function trimSize(fileSize) {
    var periodPosition = fileSize.toString().indexOf(".");
    var stringLength = fileSize.toString().length;

    // String that looks like an integer
    // so we'll add '.00'
    if (periodPosition == -1) {
        return (fileSize + ".00");
    }

    // String that looks like a number ending in decimal place period
    // (Very unlikely to happen in this sample, but we'll include it anyway
    if ((stringLength - periodPosition) == 1) {
        return (fileSize + "00");
    }

    // String that looks like a number with one decimal place
    // so we'll add one trailing zero
    if ((stringLength - periodPosition) == 2) {
        return (fileSize + "0");
    }

    // String that looks like a number with two decimal places
    // so we're happy with that
    if ((stringLength - periodPosition) == 3) {
        return (fileSize);
    }

    // String that has more than two decimal places.
    // We'll simply trim the digits past the second decimal place.
    // In a real solution you might like to determine whether the second decimal place
    // should be rounded up, depending on the value of the third decimal place, but this is not really
    // the point of this sample. Simple trimming is fine for us as we're just displaying information to
    // the user about approximate file sizes :-)
    if ((stringLength - periodPosition) >= 3) {
        return (fileSize.toString().substring(0, periodPosition + 3));
    }
}

// *********************************************************
//
// PowerPoint-Add-in-JavaScript-SliceDataChunks, https://github.com/OfficeDev/Powerpoint-Add-in-JavaScript-SliceDataChunks
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************