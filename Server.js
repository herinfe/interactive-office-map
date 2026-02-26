require("dotenv").config();

const express = require('express');
const axios = require('axios');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());

// Serve files from /public folder
app.use(express.static(path.join(__dirname, 'public')));

// ✅ Homepage route (THIS FIXES "Cannot GET /")
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'Interactive Office Map.html'));
});

const tokenEndpoint = "https://login.microsoftonline.com/ac1f7d2b-c741-43f6-9893-d39b22c46953/oauth2/v2.0/token";

const params = new URLSearchParams();
params.append('client_id', process.env.CLIENT_ID);
params.append('scope', 'https://graph.microsoft.com/.default');
params.append('client_secret', process.env.CLIENT_SECRET);
params.append('grant_type', 'client_credentials');

app.get('/getAccessToken', async (req, res) => {
    try {
        const response = await axios.post(tokenEndpoint, params);
        console.log("Access Token:", response.data.access_token);
        res.json({ accessToken: response.data.access_token });
    } catch (error) {
        console.error('Error fetching access token:', error.response?.data || error.message);
        res.status(500).json({ error: 'Failed to obtain access token' });
    }
});

app.get('/checkCalendar', async (req, res) => {
    const users = ['ric_211@eab.com', 'ric_202@eab.com', 'ric_203@eab.com', 'ric_204@eab.com', 'ric_205@eab.com', 'ric_206@eab.com', 'ric_208@eab.com', 'ric_209@eab.com', 'ric_210@eab.com', 'ric_212@eab.com', 'ric_213@eab.com', 'ric_214@eab.com', 'ric_215@eab.com', 'ric_217@eab.com', 'ric_218@eab.com', 'ric_219@eab.com', 'ric_220@eab.com', 'ric_221@eab.com', 'ric_222@eab.com', 'ric_223@eab.com', 'ric_224@eab.com', 'ric_225@eab.com', 'ric_226@eab.com', 'ric_227@eab.com', 'ric_229@eab.com', 'ric_231@eab.com', 'ric_232@eab.com', 'ric_233@eab.com', 'ric_234@eab.com', 'ric_235@eab.com', 'ric_236@eab.com', 'ric_237@eab.com', 'ric_239@eab.com', 'ric_241@eab.com', 'ric_242@eab.com', 'ric_243@eab.com', 'ric_244@eab.com', 'ric_245@eab.com', 'ric_246@eab.com', 'ric_247@eab.com', 'ric_248@eab.com', 'ric_249@eab.com', 'ric_250@eab.com', 'ric_251@eab.com', 'ric_253@eab.com', 'ric_254@eab.com', 'ric_255@eab.com', 'ric_256@eab.com', 'ric_257@eab.com', 'ric_258@eab.com', 'ric_259@eab.com', 'ric_201@eab.com'];

    try {
        const tokenResponse = await axios.post(tokenEndpoint, params);
        const accessToken = tokenResponse.data.access_token;
        const now = new Date().toISOString();

        const calendarResponses = await Promise.all(users.map(user => {
            return axios.get(`https://graph.microsoft.com/v1.0/users/${user}/calendarView?startDateTime=${now}&endDateTime=${now}`, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });
        }));

        const areBusy = calendarResponses.map(response => response.data.value.length > 0);
        const result = {};
        users.forEach((user, index) => {
            result[user] = areBusy[index];
        });

        res.json(result);
    } catch (error) {
        console.error('Error:', error.response ? error.response.data : error.message);
        res.status(500).json({ error: 'Failed to check calendar' });
    }
});

// Modified endpoint for creating an event in a user's Outlook calendar
app.post('/createEvent', async (req, res) => {
    const { userEmail, subject, start, end, attendees, bodyContent } = req.body;

    try {
        // Obtain an access token
        const tokenResponse = await axios.post(tokenEndpoint, params);
        const accessToken = tokenResponse.data.access_token;

        // Check for conflicting events
        const startDateTime = new Date(start).toISOString();
        const endDateTime = new Date(end).toISOString();
        const checkEventUrl = `https://graph.microsoft.com/v1.0/users/${userEmail}/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}`;
        const checkEventResponse = await axios.get(checkEventUrl, {
            headers: { Authorization: `Bearer ${accessToken}` }
        });

        if (checkEventResponse.data.value.length > 0) {
            // Conflicting event exists
            return res.status(400).json({ error: 'Event conflict detected. Please choose a different time.' });
        }

        // Define the event structure
        const eventData = {
            subject: subject,
            body: {
                contentType: "HTML",
                content: bodyContent
            },
            start: {
                dateTime: start,
                timeZone: "Eastern Standard Time"
            },
            end: {
                dateTime: end,
                timeZone: "Eastern Standard Time"
            },
            attendees: attendees.map(email => ({
                emailAddress: { address: email },
                type: "required"
            }))
        };

        // Create the event by making a POST request to Microsoft Graph API
        const createEventResponse = await axios.post(`https://graph.microsoft.com/v1.0/users/${userEmail}/events`, eventData, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        // Respond with the created event details
        res.json({ message: 'Event created successfully', eventData: createEventResponse.data });
    } catch (error) {
        console.error('Error creating event:', error.response ? error.response.data : error.message);
        res.status(500).json({ error: 'Failed to create event', details: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Server is running at http://localhost:${PORT}`);
});
