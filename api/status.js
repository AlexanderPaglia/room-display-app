// /api/status.js
// Using built-in fetch instead of axios to avoid dependency issues

// Configuration will be read from Environment Variables on Vercel
const CONFIG = {
    tenantId: process.env.TENANT_ID,
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    roomEmail: process.env.ROOM_EMAIL,
    scope: 'https://graph.microsoft.com/.default',
    timezone: 'America/Toronto'
};

// Cache for access token
let tokenCache = {
    token: null,
    expires: null
};

async function getAccessToken() {
    if (tokenCache.token && tokenCache.expires && new Date() < tokenCache.expires) {
        return tokenCache.token;
    }

    const tokenUrl = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', CONFIG.clientId);
    params.append('client_secret', CONFIG.clientSecret);
    params.append('scope', CONFIG.scope);
    params.append('grant_type', 'client_credentials');

    const response = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params
    });

    if (!response.ok) {
        throw new Error(`Token request failed: ${response.status}`);
    }

    const data = await response.json();
    const { access_token, expires_in } = data;
    tokenCache.token = access_token;
    tokenCache.expires = new Date(Date.now() + (expires_in - 300) * 1000);
    return access_token;
}

async function getTodayEvents() {
    const accessToken = await getAccessToken();
    const today = new Date();
    const localStartOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const localEndOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
    const startTime = localStartOfDay.toISOString();
    const endTime = localEndOfDay.toISOString();

    const graphUrl = `https://graph.microsoft.com/v1.0/users/${CONFIG.roomEmail}/calendar/calendarView?startDateTime=${startTime}&endDateTime=${endTime}&$orderby=start/dateTime&$top=50`;

    const response = await fetch(graphUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            'Prefer': `outlook.timezone="${CONFIG.timezone}"`
        }
    });

    if (!response.ok) {
        throw new Error(`Graph API request failed: ${response.status}`);
    }

    const data = await response.json();

    return (data.value || []).map(event => {
        const eventStart = new Date(event.start.dateTime);
        const eventEnd = new Date(event.end.dateTime);
        return {
            id: event.id,
            subject: event.subject || 'Meeting',
            organizer: event.organizer?.emailAddress?.name || 'Unknown',
            startTime: eventStart,
            endTime: eventEnd,
            timeRange: `${eventStart.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true, timeZone: CONFIG.timezone })} - ${eventEnd.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true, timeZone: CONFIG.timezone })}`,
        };
    });
}

export default async function handler(req, res) {
    try {
        const events = await getTodayEvents();
        const now = new Date();
        const currentEvent = events.find(event => now >= event.startTime && now < event.endTime);
        const nextEvent = events.find(event => event.startTime > now);

        // CORS headers
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

        if (req.method === 'OPTIONS') {
            return res.status(200).end();
        }

        res.status(200).json({
            success: true,
            roomEmail: CONFIG.roomEmail,
            currentTime: now.toISOString(),
            isOccupied: !!currentEvent,
            currentEvent: currentEvent || null,
            nextEvent: nextEvent || null,
            todayEventCount: events.length
        });
    } catch (error) {
        console.error('API Error:', error.message);
        res.status(500).json({
            success: false,
            error: 'Failed to fetch calendar events'
        });
    }
}
