# OutlookAddin_DataEntryEmbed
A mail task pane add-in to capture data from a user and embed this within a URL in the message 

The add-in was created using the Yeoman Generator for Office: https://dev.office.com/blogs/creating-office-add-ins-with-any-editor-introducing-yo-office

###To setup the environment to test the add-in locally
1. Install Node.js (http://nodejs.org/)
2. Open a Node.js command prompt
3. Navigate to the relevant folder
5. Run the following commands
  1. "npm install" - this will download the various dependencies 
  2. "tsd update" - this will updated the TypeScript type definition files
  3. "gulp scripts" - this will run a gulp command to build the TypeScript files
  4. "gulp serve-static" - this will run a local webserver to serve the add-in
