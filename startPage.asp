<%
' File                : startPage.asp
' Project             : PROG2001 - Web Design and Development
' Programmer          : Minchul Hwang - 8818858
' First version       : Dec. 2. 2023
' Description         : This project is a program that creates a text editor in a web browser.
'                       ASP is used as a web server, and JSON and jQuery are used to interact with the server.
'                       Also, use an additional file (getFilesList.asp) to get the list of files.


Dim fileContents
Dim existingContents
fileContents = Request.Form("textareaContents")     ' Take Contents from textarea
existingContents = ""

'Take File name
Dim fileName
fileName = Request.Form("fileName")

'Setting File path
Dim filePath
filePath = Server.MapPath("MyFiles\" & fileName & ".txt")

'Initialize JSON object to response
Dim jsonResponse
Set jsonResponse = Server.CreateObject("Scripting.Dictionary")

'Add contents in file
Dim fileObject
Set fileObject = Server.CreateObject("Scripting.FileSystemObject")
Dim getFileObject


If fileObject.FileExists(filePath) Then
    Set getFileObject = fileObject.OpenTextFile(filePath, 1)                 'If the file already exists, open it and read its contents
    
    Do Until getFileObject.AtEndOfStream                                     'Output the contents of the file with Response.Write
        Response.Write getFileObject.ReadLine() & vbCrLf
    Loop
    
    getFileObject.Close

    Set getFileObject = fileObject.OpenTextFile(filePath, 2)                 ' Write mode

    getFileObject.Write fileContents                                         ' Write the new content to the file

    getFileObject.Close                                                      ' Close the file

Else
    Set getFileObject = fileObject.CreateTextFile(filePath, True)            'Create file if it does not exist
    getFileObject.Write fileContents

    Response.Write fileContents                                              'Output the contents of the file with Response.Write

    getFileObject.Close                                                      'Close the file
End If

Set getFileObject = Nothing
Set fileObject = Nothing
%>


<!DOCTYPE html>
<html>
<head>
    <title>Text File Editor</title>
    <style>
        body {
            display: flex;
            background-color: antiquewhite;
            align-items: center;
            justify-content: center;
            height: 80vh; /* 100% 뷰포트 높이 */
            margin: 0;
        }

        form {
            text-align: center;
        }

        p {
            font-style:italic;
            font-size:50px;
            color: #361886;
        }

        div {
            font-style:inherit;
            color: #6b121b;
        }
    </style>
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
    <script type = "text/JavaScript">

        var changedTextAreaTrigger ;        // A value to check text arae is changed or not
        var callTextFileTrigger;            // A value to check textbox is changed or not
        var JSONObject;
        var contents;

        // ------- Function Comment -------
        // Name     : saveObject
        // Purpose  : When the user presses the save button, the content in the text editor is objectified using JSON, 
        //            and the textarea is saved by communicating with the server through jQuery.
        // Input    : None
        // Output   : Appropriate output according to file storage status
        // Returns  : None
        function saveObject() {
            
            var textareaValue = document.querySelector('textarea[name="textAreaContents"]').value;      // Take contents in textarea
            var fileName = document.getElementById('callTextFile').value.trim();                        // Take file name in text box
            
            if (fileName.trim() === "") {
                alert("Please enter a file name.");                                                     // if user did not put file name
                return;
            }

            JSONObject = { "contents": textareaValue };  
            var jsonString = JSON.stringify({"contents": textareaValue});
            
            // Connect with server through jQuery
            $.ajax({
                url: 'startPage.asp',  
                method: 'POST',
                data: { fileName: fileName, textareaContents: textareaValue },
                success: function (response) {
                    document.getElementById("fileStatus").innerHTML = "File Saved";
                    var selectElement = document.getElementById('textFile');
                    var newOption = document.createElement('option');
                    var existingOption = selectElement.querySelector('option[value="' + fileName + '"]');

                    if (existingOption) {                                   // if there is a file
                        existingOption.text = fileName;
                    } 
                    else {
                        var newOption = document.createElement('option');
                        newOption.value = fileName;
                        newOption.text = fileName;
                        selectElement.add(newOption);
                    }
                },
                error: function () {
                    document.getElementById("fileStatus").innerHTML = 'Error saving the file.';
                },
                complete: function () {
                    // All button are inactive condition
                    disableAllButtons();
                    callTextFileTrigger = false;
                    changedTextAreaTrigger = false;
                }
            });
        }
        
        // ------- Function Comment -------
        // Name     : saveAsObject
        // Purpose  : When the user presses the save as button, the content in the text editor is objectified using JSON, 
        //            and the textarea is saved by communicating with the server through jQuery.
        // Input    : None
        // Output   : Appropriate output according to file storage status
        // Returns  : None
        function saveAsObject() {
            var textareaValue = document.querySelector('textarea[name="textAreaContents"]').value;
            var fileName = document.getElementById('callTextFile').value.trim();
            
            if (fileName.trim() === "") {
                alert("Please enter a file name.");
                return;
            }
            else if (fileName.trim() === "Select a file") {                         // if user want to make a file which is already exist
                alert("You can not make a file named as 'Select a file'.");
                return;
            }

            JSONObject = { "contents": textareaValue };  
            var jsonString = JSON.stringify({"contents": textareaValue});

            document.getElementById("saveButton").disabled = true;
            $.ajax({
                url: 'startPage.asp',  
                method: 'POST',
                data: { fileName: fileName, textareaContents: textareaValue },
                success: function (response) {                                      // if user make a new file, the select list is updated
                    document.getElementById("fileStatus").innerHTML = "File Saved";
                    var selectElement = document.getElementById('textFile');
                    var newOption = document.createElement('option');
                    
                    newOption.value = callTextFile.value;
                    newOption.text = callTextFile.value;
                    selectElement.add(newOption);
                },
                error: function () {
                    document.getElementById("fileStatus").innerHTML = 'Error saving the file.';
                },
                complete: function () {
                    // All button are inactive condition
                    disableAllButtons();
                    callTextFileTrigger = false;
                    changedTextAreaTrigger = false;
                }
            });
        }

        // ------- Function Comment -------
        // Name     : loadSelectedFile
        // Purpose  : This function is executed when the user selects a file in select.
        //            Also, in order to maintain the contents of the file, back up the contents again with the file name intact. 
        // Input    : None
        // Output   : Appropriate output according to file storage status
        // Returns  : None
        function loadSelectedFile() {
            var textareaValue = document.querySelector('textarea[name="textAreaContents"]').value;
            var fileName = document.getElementById('callTextFile').value.trim();
            JSONObject = { "contents": textareaValue };  
            var jsonString = JSON.stringify({"contents": textareaValue});

            $.ajax({                                    // take a file and back up again
                url: 'startPage.asp',  
                method: 'POST',
                data: { fileName: fileName, textareaContents: textareaValue },
                success: function (response) {
                    document.getElementById("fileStatus").innerHTML = "File loaded";
                    var selectElement = document.getElementById('textFile');
                    var newOption = document.createElement('option');
                    var existingOption = selectElement.querySelector('option[value="' + fileName + '"]');
                    existingOption.text = fileName;
                },
                error: function () {
                        document.getElementById("fileStatus").innerHTML = 'Error saving the file.';
                        
                },
                complete: function () {
                    // All button are inactive condition
                    disableAllButtons();
                    callTextFileTrigger = false;
                    changedTextAreaTrigger = false;
                }
            });
        }

        // ------- Function Comment -------
        // Name     : updateCallTextFile
        // Purpose  : This function works when there is input or change in the textbox.
        //            Also, when a file is selected in select, the contents are retrieved and output to the textarea.
        // Input    : None
        // Output   : Appropriate output according to file storage status
        // Returns  : None
        function updateCallTextFile() {
            var selectedOption = document.getElementById('textFile');
            var callTextFile = document.getElementById('callTextFile');
            var selectedText = selectedOption.options[selectedOption.selectedIndex].text;

            callTextFile.value = selectedText;
            
            if (selectedText !== "Select a file") {                 // Get the contents of the selected file and display it in the textarea
                getFileContents(selectedText);
            } else {                                                // Clear textarea if no file is selected
                document.querySelector('textarea[name="textAreaContents"]').value = "";
            }
            // All button are inactive condition
            disableAllButtons();
            callTextFileTrigger = false;
            changedTextAreaTrigger = false;
        }
        
        // ------- Function Comment -------
        // Name     : getFileContents
        // Purpose  : This program is responsible for retrieving the contents of the file.
        // Input    : fileName      a string which is file name
        // Output   : Appropriate output according to file storage status
        // Returns  : None
        function getFileContents(fileName) {
            $.ajax({
                url: 'startPage.asp',  
                method: 'POST',
                data: { fileName: fileName },
                success: function (response) {
                    var textContent;
                    textContent = RemoveResponse(response);
                    document.querySelector('textarea[name="textAreaContents"]').value = textContent;
                    document.getElementById("fileStatus").innerHTML = "File Loaded";
                    changedTextArea();
                    loadSelectedFile();
                },
                error: function () {
                    document.getElementById("fileStatus").innerHTML = 'Error loading the file.';
                },
                complete: function () {
                    // All button are inactive condition
                    disableAllButtons();
                    callTextFileTrigger = false;
                    changedTextAreaTrigger = false;
                }
            });
        }

        // ------- Function Comment -------
        // Name     : RemoveResponse
        // Purpose  : This function cuts the code part of the contents of the file being imported.
        // Input    : response              a string which has contents
        // Output   : none
        // Returns  : slicedResponse        a string which is cut of code
        function RemoveResponse(response){
            var cutIndex = response.indexOf("\n<!DOCTYPE html>");       // Check until <!DOCTYPE html> appears.
            var slicedResponse;
            if (cutIndex !== -1) {
                slicedResponse = response.substring(0, cutIndex-3);
            } 

            return slicedResponse;
        }

        // ------- Function Comment -------
        // Name     : loadFileNames
        // Purpose  : This function is responsible for retrieving the files in the list. This utilizes getFileList.asp.
        //            It also serves to insert the file name into the option.
        // Input    : none
        // Output   : Appropriate output according to file storage status
        // Returns  : none
        function loadFileNames() {
            $.ajax({
                url: 'getFilesList.asp', 
                method: 'POST',
                success: function (response) {
                    var fileNames = JSON.parse(response);

                    var selectElement = document.getElementById('textFile');                // Clear existing options and add "Select a file" option
                    selectElement.innerHTML = '<option>Select a file</option>';             // Add option
                    
                    for (var i = 0; i < fileNames.length; i++) {
                        var newOption = document.createElement('option');
                        newOption.value = fileNames[i];
                        newOption.text = fileNames[i];
                        selectElement.add(newOption);
                    }
                },
                error: function () {
                }
            });
        }

        // ------- Function Comment -------
        // Name     : window.onload
        // Purpose  : This is used to disable all buttons when Windows is updated.
        // Input    : none
        // Output   : none
        // Returns  : none
        window.onload = function () {
            disableAllButtons();
        };

        // ------- Function Comment -------
        // Name     : disableAllButtons
        // Purpose  : A function that disables both the save and save as buttons.
        // Input    : none
        // Output   : none
        // Returns  : none
        function disableAllButtons() {
            document.getElementById("saveButton").disabled = true;
            document.getElementById("saveAsButton").disabled = true;
        }

        // ------- Function Comment -------
        // Name     : enableSaveButton
        // Purpose  : A Function to activate the save button and disable the save as button.
        // Input    : none
        // Output   : none
        // Returns  : none
        function enableSaveButton(){
            document.getElementById("saveButton").disabled = false;
            document.getElementById("saveAsButton").disabled = true;
        }

        // ------- Function Comment -------
        // Name     : enableSaveButton
        // Purpose  : A Function to activate the save as button and disable the save button.
        // Input    : none
        // Output   : none
        // Returns  : none
        function enableSaveAsButton() {
            document.getElementById("saveButton").disabled = true;
            document.getElementById("saveAsButton").disabled = false;
        }

        // ------- Function Comment -------
        // Name     : changedTextArea
        // Purpose  : A function that checks whether the textarea changes or not, and varies the method of activating the button accordingly.
        // Input    : none
        // Output   : none
        // Returns  : none
        function changedTextArea(){
            changedTextAreaTrigger = true;
            document.getElementById("fileStatus").innerHTML = "";
            var textareaValue = document.querySelector('textarea[name="textAreaContents"]').value;
            charCount = textareaValue.length;
            document.getElementById("textCount").innerHTML = "Character Count : " + charCount;
            if(callTextFileTrigger == false){
                enableSaveButton();
            }
            else{
                enableSaveAsButton();
            }
        }

        // ------- Function Comment -------
        // Name     : changedTextArea
        // Purpose  : A function that checks whether the contents of the textbox change or not and activates the corresponding button.
        // Input    : none
        // Output   : none
        // Returns  : none
        function changeTextFile(){
            callTextFileTrigger = true;
            document.getElementById("fileStatus").innerHTML = "";
            var textBoxValue = document.getElementById('callTextFile').value.trim();
            var selectedOption = document.getElementById('textFile');
            var selectedText = selectedOption.options[selectedOption.selectedIndex].text;

            if (textBoxValue === selectedText) {
                disableAllButtons(); 
            } else {
                enableSaveAsButton(); 
            }
        }

        // Function to put files in MyFiles into a list.
        loadFileNames();
    </script>
</head>

<body>
    <form name="GetContents" text>
        <p>TEXT EDITOR</p>
        <hr>
        <div>
            <input type="button" id="saveButton" onclick="saveObject();" value="Save" style="font-style: oblique;"/>
            <input type="button" id="saveAsButton" onclick="saveAsObject();" value="Save As" style="font-style: oblique;"/>
            <select id="textFile" oninput="updateCallTextFile();">
                <option>Select a file</option>
            </select>
            <input type="text" id="callTextFile" oninput="changeTextFile();"/>
            
            <div id="fileStatus"></div>
            <br>
            <textarea style="width: 500px; height: 300px; resize: none; border:10pt solid #000; border-style: double;" name="textAreaContents" oninput="changedTextArea();"></textarea>
        </div>
        <span id="textCount" style="font-style: oblique; color: rgb(119, 8, 152);"></span>
    </form>
</body>
</html>