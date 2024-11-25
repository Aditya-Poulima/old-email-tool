const { google } = require('googleapis');
const express = require('express');
const xlsx = require('xlsx');
const emailValidator = require('email-validator');
const { SMTPClient } = require('smtp-client');
const dns = require('dns');
require('dotenv').config();

const app = express();

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;
const oauth2Client = new google.auth.OAuth2(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI);

// Get valid email rows only
function getRecipientsFromExcel(filePath) {
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const recipients = xlsx.utils.sheet_to_json(worksheet, { header: 1 })
        .slice(1) // Skip header
        .filter(row => row[1] && emailValidator.validate(row[1])) // Filter for valid emails
        .map(row => ({
            name: row[0] || "Valued Client", // Use default name if empty
            email: row[1]
        }));

    return recipients;
}

async function getMXServer(domain) {
    return new Promise((resolve, reject) => {
        dns.resolveMx(domain, (err, addresses) => {
            if (err || !addresses || addresses.length === 0) {
                reject(new Error(`MX lookup failed for ${domain}`));
            } else {
                const sortedAddresses = addresses.sort((a, b) => a.priority - b.priority);
                resolve(sortedAddresses[0].exchange);
            }
        });
    });
}

async function verifyEmailSMTP(email) {
    const domain = email.split('@')[1];

    try {
        const mxServer = await getMXServer(domain);
        const client = new SMTPClient({ host: mxServer, port: 25, tls: false });

        await client.connect();
        await client.greet({ hostname: 'localhost' });
        await client.mail({ from: 'kalyan@poulimainfo.tech' });
        await client.rcpt({ to: email });
        await client.quit();

        return true; // Email exists
    } catch (error) {
        console.log(`SMTP check failed for ${email}: ${error.message}`);
        return false; // Email rejected or doesn't exist
    }
}

async function sendEmailCampaign(gmail, recipients) {
    let count = 0;

    for (const recipient of recipients) {
        try {
            const isValid = await verifyEmailSMTP(recipient.email);
            if (!isValid) {
                console.log(`Skipping invalid email: ${recipient.email}`);
                continue;
            }

            const mailOptions = {
                    from: '"Mounika" <mounika.aeleshwaram@poulimainfo.tech>',
                    to: recipient.email,
                    cc: 'saheli.samanta@poulimainfo.tech,contact@poulimainfo.tech,', // Add CC emails here
                    subject: `Following Up: Hire skilled Full Stack developers at affordable prices!`,
                    text: `Hi ${recipient.name}, just following up on my previous email regarding skilled developers. Please refer to the quoted message below.`, // Text version for non-HTML email clients
                    html: `
                    <p>Hi ${recipient.name},</p>
                    <p>I hope this message finds you well! I wanted to follow up on my previous email about our skilled Full Stack developers available at Poulima Infotech.</p>
                    <p>If you’re interested in exploring further, please feel free to reach out. I’d be happy to discuss your requirements and share relevant developer profiles with you.</p>
                    <p>Looking forward to your response!</p>
               
                    <p>Thanks & Regards,</p>
               
                    <table style="width: 100%; max-width: 500px; border-collapse: collapse;">
                        <tr style="height: 100px;">
                            <td style="padding: 10px; width: 42%; vertical-align: middle; text-align: center; box-shadow: 10px 0 10px -10px rgba(0, 0, 0, 0.3);">
                                <img src="https://www.poulimainfo.tech/wp-content/uploads/2022/04/Poulima-Infotech-Color-Logo.png" alt="logo" style="width: 100%; max-width: 150px; height: auto;" />
                            </td>
                            <td style="padding: 10px; width: 58%; vertical-align: middle; font-family: Arial, sans-serif;">
                                <p style="font-weight: bold; margin: 0;">Mounika Aeleshwaram</p>
                                <p style="font-weight: bold; margin: 0;">Business Development Executive</p>
                                <p style="margin: 0;"><b>T:</b> +91 (912) 120-2538</p>
                                <p style="margin: 0;"><b>E:</b> <a href="mailto:mounika.aeleshwaram@poulimainfo.tech">mounika.aeleshwaram@poulimainfo.tech</a> | <a href="https://www.poulimainfo.tech/">www.poulimainfo.tech</a></p>
                                <p style="margin: 0;">Whitefield, Bangalore | 560066, India</p>
                            </td>
                        </tr>
                    </table>
               
                    <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
                   
                    <p style="color: #500050;"><strong>On Sun, 10 Nov 2024 at 20:00, Mounika &lt;mounika.aeleshwaram@poulimainfo.tech&gt; wrote:</strong></p>
                   
                    <blockquote style="border-left: 1px solid #CCCCCC; padding-left: 1ex; color: #500050;">
                        <p>Hi Team,</p>
                        <p>Greetings! My name is Mounika, and I represent the On-Demand Software Developers Team at Poulima Infotech.</p>
                        <p>Poulima Infotech is an India-based, ISO 9001:2015 Certified company recognized by DPIIT. We offer highly skilled React, React Native developers and programmers for hire on a contract basis.</p>
                        <p>If you are interested, I would be happy to share our developer profiles with you.</p>
                        <p>Thank you for considering us. We look forward to the possibility of working together and meeting your requirements.</p>
                        <p>For instant communication, please WhatsApp us at +91 9121202538.</p>
                        <p>Looking forward to hearing from you soon!</p>
                        <p>Thanks & Regards,</p>
                       
                        <table style="width: 100%; max-width: 500px; border-collapse: collapse;">
                            <tr style="height: 100px;">
                                <td style="padding: 10px; width: 42%; vertical-align: middle; text-align: center; box-shadow: 10px 0 10px -10px rgba(0, 0, 0, 0.3);">
                                    <img src="https://www.poulimainfo.tech/wp-content/uploads/2022/04/Poulima-Infotech-Color-Logo.png" alt="logo" style="width: 100%; max-width: 150px; height: auto;" />
                                </td>
                                <td style="padding: 10px; width: 58%; vertical-align: middle; font-family: Arial, sans-serif;">
                                    <p style="font-weight: bold; margin: 0;">Mounika Aeleshwaram</p>
                                    <p style="font-weight: bold; margin: 0;">Business Development Executive</p>
                                    <p style="margin: 0;"><b>T:</b> +91 (912) 120-2538</p>
                                    <p style="margin: 0;"><b>E:</b> <a href="mailto:mounika.aeleshwaram@poulimainfo.tech">mounika.aeleshwaram@poulimainfo.tech</a> | <a href="https://www.poulimainfo.tech/">www.poulimainfo.tech</a></p>
                                    <p style="margin: 0;">Whitefield, Bangalore | 560066, India</p>
                                </td>
                            </tr>
                        </table>
                    </blockquote>
                    `
                };
               


            await gmail.users.messages.send({ userId: 'me', requestBody: { raw: Buffer.from(mailOptions).toString('base64') } });
            console.log(`${++count}/${recipients.length} emails sent to ${recipient.email}`);
        } catch (error) {
            console.log(`Error sending email to ${recipient.email}: ${error.message}`);
        }

        await delay(5000); // Delay to avoid rate limiting
    }
}

app.get('/oauth2callback', async (req, res) => {
    const code = req.query.code;

    if (code) {
        try {
            const { tokens } = await oauth2Client.getToken(code);
            oauth2Client.setCredentials(tokens);
            const gmail = google.gmail({ version: 'v1', auth: oauth2Client });

            const recipients = getRecipientsFromExcel('email-api.xlsx');
            await sendEmailCampaign(gmail, recipients);

            res.send('Email campaign completed.');
        } catch (error) {
            console.error(`Error in OAuth callback: ${error.message}`);
            res.status(500).send('An error occurred during email campaign.');
        }
    } else {
        res.status(400).send('Authorization code is missing.');
    }
});

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

app.listen(3000, () => console.log('Server is running on http://localhost:3000'));


