# student-tracker-node

### Setup
Becuase the app is intended to be run from the command line by a single user, the only way to authenticate with google is to create credentials in the Google Cloud Platform Developer Console or receive a credentials.json file from the developer (jrk.123.jk@gmail.com).  These steps will assume a credentials.json file has been created and exists in the project directory.  If you have not created one, follow instruction for [Creating a project & enabling the API](https://developers.google.com/workspace/guides/create-project) and [Creating credentials](https://developers.google.com/workspace/guides/create-credentials) for a desktop application.

1. Install modules
`npm install`

2. Download the file with student data into the project directory as a csv and change the STUDENT_FILENAME variable at the top of the index.js script to match the file name

3. Run the program
`node .`

4. If this is your first time running the program in this directory, you'll be prompted to follow a link to sign in via a browser where you'll sign in, receive a code, then paste the code back into the terminal
    Note: you might get a 403 error on the sign in page that indicates you haven't been added to the list of test users.  Contact Joe at jrk.123.jk@gmail.com to be added to the list of test users.  The only way to avoid manually adding test users is to submit the app through the Google [verification process](https://support.google.com/cloud/answer/9110914).