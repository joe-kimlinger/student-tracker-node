const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const nodemailer = require("nodemailer");
const csv = require('csv-parser');
const sqlite3 = require('sqlite3').verbose();

// If modifying these scopes, delete token.json.
const SCOPES = ['https://mail.google.com/'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';
const STUDENT_FILENAME = 'Work Assessment Final List Template - Sheet1.csv'

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
	if (err) return console.log('Error loading client secret file:', err);
	// Autorize using Gmail, parse the student data file, and setup the database connection
    authorize(JSON.parse(content)).then((auth) => {
        Promise.all([
            readStudentFile(STUDENT_FILENAME),
            setupDatabaseConnection()
        ]).then((results) => {
            students = results[0]
            dbConn = results[1]

            // Check to make sure that sending email to current students does not exceed daily email usage limit and
            // setup NodeMailer connection
            Promise.all([
                checkEmailLimits(dbConn, students),
                setupNodeMailerConnection(auth)
            ]).then((res) => {
                console.log("Attempting to send " + students.length + " emails.  " + res[0] + " emails sent so far today.")
                sendEmails(transporter, students, dbConn)

            }).catch((err) => console.log(err));

        }).catch((err) => console.log(err));
    }).catch((err) => console.log(err));
});

/**
 * Create an OAuth2 client with the given credentials
 * 
 * @param {Object} credentials The authorization client credentials.
 */
function authorize(credentials) {
	const {client_secret, client_id, redirect_uris} = credentials.installed;
	const oAuth2Client = new google.auth.OAuth2(
			client_id, client_secret, redirect_uris[0]);

	// Check if we have previously stored a token.
	return new Promise((resolve, reject) => {
        fs.readFile(TOKEN_PATH, (err, token) => {
            if (err) return getNewToken(oAuth2Client).then((token) => resolve(token));
            oAuth2Client.setCredentials(JSON.parse(token));
            return resolve(oAuth2Client);
        });
    });
}

/**
 * Get and store new token after prompting for user authorization, return a promise with the authorized OAuth2 client.
 * 
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 */
function getNewToken(oAuth2Client) {
	const authUrl = oAuth2Client.generateAuthUrl({
		access_type: 'offline',
		scope: SCOPES,
	});
	console.log('Authorize this app by visiting this url:', authUrl);
	const rl = readline.createInterface({
		input: process.stdin,
		output: process.stdout,
	});
	return new Promise((resolve, reject) => {
        rl.question('Enter the code from that page here: ', (code) => {
            rl.close();
            oAuth2Client.getToken(code, (err, token) => {
                if (err) return reject('Error: Error retrieving access token');
                oAuth2Client.setCredentials(token);
                // Store the token to disk for later program executions
                fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
                    if (err) return console.error(err);
                    console.log('Token stored to', TOKEN_PATH);
                });
                return resolve(oAuth2Client);
            });
        });
    });
}


/**
 * Create student objects from student file data
 *
 * @param {String} filename Name of file with students.
 */
function readStudentFile(filename) {	
    students = []
    
    return new Promise((resolve, reject) => {
        fs.createReadStream(filename)
            .on('error', () => {
                return reject("Error: There was an error reading the student data file.  Is the file in the same directory and the STUDENT_FILENAME variable updated?")
            })
            .pipe(csv())
            .on('data', (row) => {
                // Move each student onto an array to be processed by the cleanStudentData function
                students.push(row)
            })
            .on('end', () => {
                if (!students.length){
                    return reject("Error: No students found in csv.");
                }
                console.log('CSV file successfully processed.');
                
                return resolve(cleanStudentData(students))
            })
    });
}


/**
 * Parse and clean an array that comes directly from the student csv
 * 
 * @param {Array} students data read in from csv.
*/
function cleanStudentData(students) {
    // Find all the columns the contain IP in the label as these likely represent the IP scores
    ipCols = Object.keys(students[0]).filter((key) => key.trim().match(/^IP\d.*/))
    ipSettings = []

    // Parse out the total points for the IP, which exists in the column header
    ipCols.forEach((col) => {
        x = col.match(/^\s*(IP\d).*\/(\d{2})/)
        setting = {}
        setting['rawName'] = x[0]
        setting['name'] = x[1]
        setting['totalPoints'] = x[2]
        ipSettings.push(setting)
    });

    students = students.map(student => createStudentObject(student, ipSettings))

    return students;
}


/**
 * Create a new student object with the same information as the raw csv, but transformed and cleaned
 * 
 * @param {Object} student Object properties match the csv column names.
 * @param {Array} ipSettings Each element contains info about an IP
 */
function createStudentObject(student, ipSettings) {

    // Create a new student object with info from the table, but cleaned
    newStudent = {}
    newStudent['firstName'] = student['Student'].split(',')[1]
    newStudent['lastName'] = student['Student'].split(',')[0]
    newStudent['attendance'] = student['Attendance']
    newStudent['email'] = student['Email']

    // Each ip column for the current student is transformed into a new object which contains info about that student's performance in the IP
    totalPercentage = 0
    ipSettings.forEach((ip) => {
        ipScore = student[ip['rawName']]
        ipObject = {}
        ipObject['name'] = ip['name']
        ipObject['score'] = ipScore
        percentage = Math.round(ipScore / ip['totalPoints'] * 100)
        ipObject['percentage'] = percentage
        totalPercentage += percentage
        newStudent[ip['rawName']] = ipObject
    });
    newStudent['totalPercentage'] = totalPercentage / ipSettings.length


    // Create a table in html using all values from the IP objects created above
    ipTable = '<table border="1"><tr><th>IP</th><th>Score</th><th>Percentage</th></tr>'
    
    ipSettings.forEach((ip) => {
        row = '<tr align="center">'
        studentIp = newStudent[ip['rawName']]
        row += "<td>" + studentIp['name'] + "</td>"
        row += "<td>" + studentIp['score'] + "/" + ip['totalPoints'] + "</td>"
        row += "<td>" + studentIp['percentage'] + "%</td>"
        row += "</tr>"
        ipTable += row
    })
    ipTable += "</table>"
    newStudent['ipTable'] = ipTable

    return newStudent
}


/**
 * Setup a database connection using Sqlite3 file DB
 */
function setupDatabaseConnection() {	
    return new Promise((resolve, reject) => {
        db = new sqlite3.Database('./db/email-tracker.db', (err) => {
            if (err) {
                reject(err.message);
            } else {
                console.log('Connected to the email tracker database.');
                resolve(db);    
            }
        });
    });
}


 /**
 * Send emails to each user from a file with the name data.csv
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function setupNodeMailerConnection(auth) {	
    return new Promise((resolve, reject) => {
        // Get email for the current logged in user in order to initialize NodeMailer connection
        getUserEmail(auth).then((email) => {
            transporter = nodemailer.createTransport({
                host: 'smtp.gmail.com',
                port: 465,
                secure: true,
                auth: {
                    type: 'OAuth2',
                    user: email,
                    clientId: auth['_clientId'],
                    clientSecret: auth['_clientSecret'],
                    refreshToken: auth['credentials']['refresh_token'].replace(/\//, ''),
                    accessToken: auth['credentials']['access_token'],
                    expires: auth['credentials']['expiry_date']
                }
            });
            return resolve(transporter);

        }, err => reject(err));
    });
}


/**
 * Get the email for the user to be used to setup NodeMailer connection
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function getUserEmail(auth) {	
    const gmail = google.gmail({version: 'v1', auth});

    return new Promise((resolve, reject) => {
        gmail.users.getProfile({'auth': auth, 'userId': 'me'}, function(err, response) {
            if (err){
                return reject(err);
            }
            if ('data' in response && 'emailAddress' in response['data']){
                return resolve(response['data']['emailAddress'])
            } else {
                return reject("Error: Error getting user email address")
            }
        });
    })
}


/**
 * Check that the number of emails sent today does not exceed allowed limits.
 *
 * @param {sqlite3.Database} dbConn A sqlite database connection.
 * @param {Array} students Array of student objects.
 */
function checkEmailLimits(dbConn, students){

    // A log for the number of emails sent each day is contained in the EmailCountLog table w/ schema (EmailCount int, date text)
    sql = `SELECT EmailCount
            FROM EmailCountLog
            WHERE Date  = ?`;
    date = new Date(new Date().getFullYear(),new Date().getMonth() , new Date().getDate())

    // Based on limit of 500 messages per day from google https://nodemailer.com/usage/using-gmail/
    dailyLimit = 500

    if (students.length > dailyLimit)
        return Promise.reject("Error: trying to send more emails (" + students.length + ") than the daily limit (" + dailyLimit + ").")

    return new Promise((resolve, reject) => {
        dbConn.get(sql, [date], (err, row) => {
            if (err) {
                // If table doesn't exist, which occurs on the first run of the script, create it
                if (err.message.trim() == 'SQLITE_ERROR: no such table: EmailCountLog'){
                    dbConn.run('CREATE TABLE EmailCountLog(EmailCount integer, date text)');
                    return resolve(0)
                } else {
                    return reject(err.message);
                }
            }
            if (row){
                // If there is a record for today, there must have been emails sent today, so check that sending students.length more doens't exceed the limit
                if (row.EmailCount + students.length > dailyLimit){
                    errorMsg = "Error: Trying to send emails to " + students.length + " students "
                    errorMsg += "would exceed daily limit of " + dailyLimit
                    errorMsg += " (" + row.EmailCount + " emails sent so far today)."
                    return reject(errorMsg)
                } else {
                    resolve(row.EmailCount)
                }
            } else {
                return resolve(0)
            }
        });
    });
}


/**
 * Send emails to each user from a file with the name data.csv
 *
 * @param {nodemailer.Transporter<T>} transporter Authorized NodeMailer transporter.
 * @param {Array} students Array of student objects.
 * @param {sqlite3.Database} dbConn A sqlite database connection.
 */
function sendEmails(transporter, students, dbConn) {	

	fs.readFile('email-template.html', 'utf8' , (err, data) => {
        if (err) {
            console.error(err)
            return
        }

        // Replace the email template text with values from the student objects
        students.map(s => createEmailFromTemplate(s, data))

        // Send an email for each student and mark when it's been delivered
        totalEmails = students.length
        sentEmails = 0
        receivedEmails = 0
        errEmails = 0
        students.map((s) => {
            sendEmail(s, transporter).then((res, err) => {
                if (err){
                    errEmails++;
                } else if (!('accepted' in res) || res['accepted'].length < 1){
                    errEmails++;
                } else {
                    receivedEmails++;
                }
            }, (err) => {
                errEmails++;
                console.log("\nError: Error sending email to " + s['email'], err.message);
            })
            sentEmails++;
        });
        
        // Every .5 seconds, write an update to the console indicating how many emails have been sent/delivered
        consoleWriter = setInterval(() => {
            process.stdout.cursorTo(0);
            msg = "Emails sent: " + sentEmails + "/" + totalEmails + "\t"
            msg += "Emails received: " + receivedEmails + "/" + sentEmails + "\t"
            msg += "Emails not received: " + errEmails + "/" + sentEmails + "\t"
            process.stdout.write(msg);

            // If all emails have been delivered (or errored out), stop writing the status
            if (receivedEmails + errEmails == totalEmails){
                clearInterval(consoleWriter)
                console.log("\nDone sending emails.")
                saveEmailCountToDb(totalEmails, dbConn)
            }
        }, 500);
    })            
}


/**
 * Create an email using the template and the info contained in the student object.
 * 
 * @param {Object} student contains data for a single student.
 * @param {String} templateText an html template.
 */
 function createEmailFromTemplate(student, templateText) {

    // General status messages to indicate the performance of the student
    statusMessages = {
        'well': "Keep up the good work!  You have been doing well and we'd like to see you continue on this track in order to be successful in this course.",
        'poor': "Your IP scores are currently not where we'd like them to be moving into the next module.  Don't hesitate to ask for help; let's work together on improving those scores!"
    }
    attendanceMessages = {
        'perfect': "Amazing, perfect attendance!  Your dedication does not go unnoticed, keep it up!!!",
        'well': "Your dedication does not go unnoticed!  Attendance up plays a huge part in your success, so keep showing up!",
        'poor': "Your attendance is low.  Make sure you're present to get the information you need to succeed!"
    }
    templateText = templateText.replace(/{{DATE}}/g, (new Date()).toLocaleDateString("en-US"))
    templateText = templateText.replace(/{{STUDENT_FIRST_NAME}}/g, student['firstName'])
    templateText = templateText.replace(/{{STUDENT_LAST_NAME}}/g, student['lastName'])

    // Based on analysis of spreadsheet, I determined there was a threshold at 75% that determined if students are doing well (recommended) vs. poorly (not recommended)
    // Here, I set the status message based on this threshold of 75%
    statusMessage = statusMessages['well']
    if (student['totalPercentage'] < '75'){
        statusMessage = statusMessages['poor']
    }
    templateText = templateText.replace(/{{STATUS_MSG}}/g, statusMessage)

    templateText = templateText.replace(/{{TOTAL_PERCENTAGE}}/g, student['totalPercentage'])

    templateText = templateText.replace(/{{IP_TABLE}}/g, student['ipTable'])
    templateText = templateText.replace(/{{ATTENDANCE}}/g, student['attendance'])

    // Set attendance message based on threshold of 95%, which I determinted to be a marker of poor attendance 
    attendanceMessage = attendanceMessages['well']
    if (student['attendance'] == '100'){
        attendanceMessage = attendanceMessages['perfect']
    } else if (student['attendance'] <= '95') {
        attendanceMessage = attendanceMessages['poor']
    }
    templateText = templateText.replace(/{{ATTENDANCE_MSG}}/g, attendanceMessage)

    student['emailText'] = templateText
}

/**
 * Send an email using the gmail API
 * 
 * @param {Object} student contains data for a single student.
 * @param {nodemailer.Transporter<T>} auth An authorized OAuth2 client.
 */
function sendEmail(student, transporter){

    emailSubject = (new Date()).toLocaleDateString("en-US") + " Status report for " + student['firstName'] + " " + student['lastName']
    return transporter.sendMail({
        to: student['email'],
        subject: emailSubject,
        html: student['emailText'],
        auth: {
            user: transporter.transporter.auth.user,
            refreshToken: transporter.options.auth.refreshToken,
            accessToken: transporter.options.auth.accessToken,
            expires: transporter.options.auth.expires
        }
    });
}

/**
 * Update the daily count to reflect the number of emails sent
 * 
 * @param {number} numEmails Number of emails sent in this run of the program.
 * @param {sqlite3.Database} dbConn A sqlite database connection.
 */
function saveEmailCountToDb(numEmails, dbConn){
    date = new Date(new Date().getFullYear(),new Date().getMonth() , new Date().getDate())
    sql = `Update EmailCountLog
           set EmailCount = EmailCount + ?
           where Date = ?`
    
    dbConn.run(sql, [numEmails, date], function(err){
        if (err){
            return console.log("Error updating database with new daily count", err.message)
        }
        // If the update was not performed, likely the row for today doesn't exist yet, so perform an insert instead
        if (this.changes < 1){
            sql = `Insert into EmailCountLog
                   (EmailCount, Date) values (?, ?)`
            dbConn.run(sql, [numEmails, date], function(err){
                if (err){
                    return console.log("Error updating database with new daily count", err.message)
                }
                console.log("Updated database with new daily email count.");
            });
        } else {
            console.log("Updated database with new daily email count.");
        }
    });
}