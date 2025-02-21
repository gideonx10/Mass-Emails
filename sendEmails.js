require('dotenv').config();
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const fs = require('fs');

// Load the Excel file
const workbook = xlsx.readFile('students.xlsx'); // Change filename if needed
const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
const sheet = workbook.Sheets[sheetName];

// Convert sheet to JSON
const students = xlsx.utils.sheet_to_json(sheet);

// Extract emails
const emails = students.map(student => student.Email).filter(email => email); // Assuming column name is 'Email'

console.log(`Found ${emails.length} emails.`);

// Configure Nodemailer Transporter
const transporter = nodemailer.createTransport({
    service: 'gmail', // You can change this based on your provider
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
    },
});

// Email Sending Function
const sendEmail = async (email) => {
    try {
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: email,
            subject: 'Your Subject Here',
            text: 'Hello, this is a test email from Nodemailer!',
        };

        await transporter.sendMail(mailOptions);
        console.log(`Email sent to ${email}`);
    } catch (error) {
        console.error(`Failed to send email to ${email}: ${error.message}`);
    }
};

// Send emails one by one
(async () => {
    for (const email of emails) {
        await sendEmail(email);
    }
    console.log('All emails sent!');
})();