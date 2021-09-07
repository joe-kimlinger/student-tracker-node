const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const nodemailer = require("nodemailer");
const csv = require('csv-parser');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.send'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
	if (err) return console.log('Error loading client secret file:', err);
	// Authorize a client with credentials, then call the Gmail API.
	authorize(JSON.parse(content), sendEmails);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
	const {client_secret, client_id, redirect_uris} = credentials.installed;
	const oAuth2Client = new google.auth.OAuth2(
			client_id, client_secret, redirect_uris[0]);

	// Check if we have previously stored a token.
	fs.readFile(TOKEN_PATH, (err, token) => {
		if (err) return getNewToken(oAuth2Client, callback);
		oAuth2Client.setCredentials(JSON.parse(token));
		callback(oAuth2Client);
	});
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
	const authUrl = oAuth2Client.generateAuthUrl({
		access_type: 'offline',
		scope: SCOPES,
	});
	console.log('Authorize this app by visiting this url:', authUrl);
	const rl = readline.createInterface({
		input: process.stdin,
		output: process.stdout,
	});
	rl.question('Enter the code from that page here: ', (code) => {
		rl.close();
		oAuth2Client.getToken(code, (err, token) => {
			if (err) return console.error('Error retrieving access token', err);
			oAuth2Client.setCredentials(token);
			// Store the token to disk for later program executions
			fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
				if (err) return console.error(err);
				console.log('Token stored to', TOKEN_PATH);
			});
			callback(oAuth2Client);
		});
	});
}

/**
 * Send emails to each user from a file with the name data.csv
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function sendEmails(auth) {	
    students = []

	return fs.createReadStream('test.csv')
		.pipe(csv())
		.on('data', (row) => {
            students.push(row)
	})
		.on('end', () => {
			console.log('CSV file successfully processed');
            if (!students.length){
                console.log("No students found in csv");
                return;
            }

            students = parseStudentFile(students)

            fs.readFile('email-template.html', 'utf8' , (err, data) => {
                if (err) {
                    console.error(err)
                    return
                }

                students.map(s => createEmailFromTemplate(s, data))
                students.map(s => sendEmail(s, auth))
            })            
	});
}

/**
 * Parse and clean an array that comes directly from the student csv
 * @param {Array} students data read in from csv.
 */
function parseStudentFile(students) {
    ipCols = Object.keys(students[0]).filter((key) => key.trim().match(/^IP\d.*/))
    ipSettings = []
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
 * Create a new student object with the same information as the raw csv, 
 * but transformed and cleaned
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


    // Create a table in html using all of the values from the IP scores
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
 * Create an email using the template and the info contained in the 
 * student object.
 * @param {Object} student contains data for a single student.
 * @param {String} templateText an html template.
 */
 function createEmailFromTemplate(student, templateText) {

    statusMessages = {
        'well': "Keep up the good work!  You have been doing well and we'd like to see you continue on this track in order to be successful in this course.",
        'poor': "Your IP scores are currently not where we'd like them to be moving into the next module.  Don't hesitate to ask for help; let's work together on improving those scores!"
    }
    attendanceMessages = {
        'perfect': "Amazing, perfect attendance!  Your dedication does not go unnoticed, keep it up!!!",
        'well': "Your dedication does not go unnoticed!  Attendance up plays a huge part in your success, so keep showing up!",
        'poor': "Your attendance is low.  Make sure you're present to get the information you need to succeed!"
    }
    options = {year: 'numeric', month: 'long', day: 'numeric' };
    date = (new Date()).toLocaleDateString("en-US")
    templateText = templateText.replace(/{{DATE}}/g, date)
    templateText = templateText.replace(/{{STUDENT_FIRST_NAME}}/g, student['firstName'])
    templateText = templateText.replace(/{{STUDENT_LAST_NAME}}/g, student['lastName'])

    // Set status message based on threshold of 75%
    statusMessage = statusMessages['well']
    if (student['totalPercentage']){
        statusMessage = statusMessages['poor']
    }
    templateText = templateText.replace(/{{STATUS_MSG}}/g, statusMessage)

    templateText = templateText.replace(/{{TOTAL_PERCENTAGE}}/g, student['totalPercentage'])

    templateText = templateText.replace(/{{IP_TABLE}}/g, student['ipTable'])
    templateText = templateText.replace(/{{ATTENDANCE}}/g, student['attendance'])

    // Set attendance message based on threshold of 95%
    attendanceMessage = attendanceMessages['well']
    if (student['attendance'] == '100'){
        attendanceMessage = attendanceMessages['perfect']
    } else if (student['attendance'] <= '95') {
        attendanceMessage = attendanceMessages['poor']
    }
    templateText = templateText.replace(/{{ATTENDANCE_MSG}}/g, attendanceMessage)

    student['emailText'] = templateText

    var str = ["Content-Type: text/html; charset=\"UTF-8\"\n",
        "MIME-Version: 1.0\n",
        "Content-Transfer-Encoding: 7bit\n",
        "to: ", student['email'], "\n",
        "from: ", 'me', "\n",
        "subject: ", date + " Status report for " + student['firstName'] + " " + student['lastName'] , "\n\n",
        student['emailText']
    ].join('');

    var raw = new Buffer.from(str).toString("base64").replace(/\+/g, '-').replace(/\//g, '_');
    student['rawEmail'] = raw
}

/**
 * Send an email using the gmail API
 * @param {Object} student contains data for a single student.
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function sendEmail(student, auth){
    const gmail = google.gmail({version: 'v1', auth});

    gmail.users.messages.send({
        auth: auth,
        userId: 'me',
        resource: {
            raw: student['rawEmail']
        }
    }, function(err, response) {
        if (err){
            console.log("An error occurred sending the email to " + student['firstName'] + " " + student['lastName'])
            console.log(err['message'])
        }
    });
}
