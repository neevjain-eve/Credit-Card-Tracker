// Pulls the master transactions/submissions file from OneDrive using
// app-only Microsoft Graph credentials (client credentials flow — no
// interactive browser login involved), then splits it per client and
// writes clients/<safe-username>.json files into the repo checkout.
//
// Required environment variables (set as GitHub Actions secrets):
//   ONEDRIVE_TENANT_ID     - Azure AD tenant ID (GUID) or verified domain
//   ONEDRIVE_CLIENT_ID     - Azure app registration's Application (client) ID
//   ONEDRIVE_CLIENT_SECRET - a client secret created for that app registration
//   ONEDRIVE_SHARE_LINK    - the OneDrive "anyone with the link can edit"
//                            share URL for the master data file
//
// The Azure app registration needs an *Application* permission (not
// Delegated) of Files.Read.All (or Sites.Read.All if the file lives in
// SharePoint document libraries), with admin consent granted in Azure
// Portal. This is separate from the app's existing interactive
// (delegated) OneDrive login used by the "Connect OneDrive" button.

const fs = require('fs');
const path = require('path');

const TENANT_ID = process.env.ONEDRIVE_TENANT_ID;
const CLIENT_ID = process.env.ONEDRIVE_CLIENT_ID;
const CLIENT_SECRET = process.env.ONEDRIVE_CLIENT_SECRET;
const SHARE_LINK = process.env.ONEDRIVE_SHARE_LINK;

function cssSafeId(str) {
    return String(str).replace(/[^a-zA-Z0-9]/g, '_');
}

function encodeShareUrl(url) {
    const base64 = Buffer.from(url, 'utf8').toString('base64')
        .replace(/=+$/, '')
        .replace(/\//g, '_')
        .replace(/\+/g, '-');
    return 'u!' + base64;
}

async function getAppOnlyToken() {
    const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const body = new URLSearchParams({
        client_id: CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
    });
    const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body
    });
    if (!res.ok) {
        throw new Error(`Failed to get app-only token: ${res.status} ${await res.text()}`);
    }
    const data = await res.json();
    return data.access_token;
}

async function resolveSharedItem(token) {
    const encoded = encodeShareUrl(SHARE_LINK);
    const res = await fetch(`https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem?$select=id,parentReference`, {
        headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
        throw new Error(`Failed to resolve shared OneDrive item: ${res.status} ${await res.text()}`);
    }
    const item = await res.json();
    return { driveId: item.parentReference.driveId, itemId: item.id };
}

async function downloadItemContent(token, ref) {
    const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${ref.driveId}/items/${ref.itemId}/content`, {
        headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
        throw new Error(`Failed to download OneDrive file content: ${res.status} ${await res.text()}`);
    }
    return res.json();
}

function buildClientPayload(allData, clientName) {
    const norm = s => String(s || '').trim().toLowerCase();
    const transactions = (allData.transactions || []).filter(t => norm(t.name) === norm(clientName));
    const relevantIds = new Set(transactions.map(t => t.submissionId).filter(Boolean));
    const submissions = (allData.submissions || []).filter(s => relevantIds.has(s.submissionId));
    return { transactions, submissions, exportedAt: Date.now() };
}

async function main() {
    if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !SHARE_LINK) {
        console.error('Missing one or more required env vars: ONEDRIVE_TENANT_ID, ONEDRIVE_CLIENT_ID, ONEDRIVE_CLIENT_SECRET, ONEDRIVE_SHARE_LINK');
        process.exit(1);
    }

    const clientMapPath = path.join(__dirname, '..', '..', 'client-map.json');
    const clients = JSON.parse(fs.readFileSync(clientMapPath, 'utf8'));

    console.log('Getting app-only Graph token...');
    const token = await getAppOnlyToken();

    console.log('Resolving shared OneDrive item...');
    const ref = await resolveSharedItem(token);

    console.log('Downloading master data file...');
    const allData = await downloadItemContent(token, ref);

    const outDir = path.join(__dirname, '..', '..', 'clients');
    fs.mkdirSync(outDir, { recursive: true });

    for (const client of clients) {
        if (client.role === 'Admin') continue;
        const payload = buildClientPayload(allData, client.name);
        const filePath = path.join(outDir, `${cssSafeId(client.username)}.json`);
        fs.writeFileSync(filePath, JSON.stringify(payload, null, 2));
        console.log(`Wrote ${filePath}: ${payload.transactions.length} transaction(s) for ${client.name}`);
    }

    console.log('Done.');
}

main().catch(err => {
    console.error(err);
    process.exit(1);
});
