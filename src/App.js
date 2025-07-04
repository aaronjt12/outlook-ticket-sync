import React, { useState, useEffect } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";

const msalConfig = {
  auth: {
    clientId: "5c1e64c0-76f2-4200-8ee5-b3b3d19b53da",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
};
const msalInstance = new PublicClientApplication(msalConfig);

const loginRequestOutlook = {
  scopes: ["openid", "profile", "User.Read", "Mail.Read"],
  prompt: "select_account"
};
const loginRequestSharePoint = {
  scopes: ["openid", "profile", "Sites.ReadWrite.All"],
  prompt: "select_account"
};

function DualLogin({
  useSameAccount,
  setUseSameAccount,
  outlookToken,
  setOutlookToken,
  sharepointToken,
  setSharepointToken,
  setOutlookAccount,
  setSharepointAccount,
}) {
  const { instance } = useMsal();

  const handleLoginOutlook = async () => {
    const res = await instance.loginPopup(loginRequestOutlook);
    setOutlookToken(res.accessToken);
    setOutlookAccount(res.account);
  };
  const handleLoginSharePoint = async () => {
    const res = await instance.loginPopup(loginRequestSharePoint);
    setSharepointToken(res.accessToken);
    setSharepointAccount(res.account);
  };
  const handleLogout = () => {
    instance.logoutPopup();
    setOutlookToken(null);
    setSharepointToken(null);
    setOutlookAccount(null);
    setSharepointAccount(null);
  };
  return (
    <div style={{ marginBottom: 20 }}>
      <label>
        <input
          type="checkbox"
          checked={useSameAccount}
          onChange={e => setUseSameAccount(e.target.checked)}
        />
        Use the same account for both Outlook and SharePoint
      </label>
      <div style={{ marginTop: 10 }}>
        {useSameAccount ? (
          <button onClick={handleLoginOutlook}>Login with Microsoft</button>
        ) : (
          <>
            <button onClick={handleLoginOutlook} style={{ marginRight: 8 }}>
              Login to Outlook
            </button>
            <button onClick={handleLoginSharePoint}>Login to SharePoint</button>
          </>
        )}
        <button onClick={handleLogout} style={{ marginLeft: 16 }}>Logout</button>
      </div>
    </div>
  );
}

function EmailList({ token, selectedEmails, setSelectedEmails }) {
  const [emails, setEmails] = useState([]);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!token) return;
    setLoading(true);
    fetch("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20", {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then((res) => res.json())
      .then((data) => {
        setEmails(data.value || []);
        setLoading(false);
      });
  }, [token]);

  const toggleEmail = (id) => {
    setSelectedEmails((prev) =>
      prev.includes(id) ? prev.filter((eid) => eid !== id) : [...prev, id]
    );
  };

  return (
    <div>
      <h3>Select Emails</h3>
      {loading && <div>Loading emails...</div>}
      <ul style={{ maxHeight: 300, overflowY: "auto", padding: 0 }}>
        {emails.map((email) => (
          <li key={email.id} style={{ listStyle: "none", marginBottom: 8 }}>
            <label>
              <input
                type="checkbox"
                checked={selectedEmails.includes(email.id)}
                onChange={() => toggleEmail(email.id)}
              />
              <b>{email.subject || "(No Subject)"}</b> — {email.from?.emailAddress?.address}
            </label>
          </li>
        ))}
      </ul>
    </div>
  );
}

function FieldMapping({ columns, fieldMapping, setFieldMapping }) {
  // Only show columns that are not hidden and are not system fields
  const selectableColumns = columns.filter(col => !col.hidden && !col.readOnly && col.name !== 'ContentType' && col.name !== 'Attachments');
  return (
    <div style={{ margin: '20px 0', padding: 10, border: '1px solid #ccc', borderRadius: 6 }}>
      <h4>SharePoint Field Mapping</h4>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {['subject', 'description', 'user', 'ticketnumber'].map(appField => (
          <label key={appField}>
            {appField.charAt(0).toUpperCase() + appField.slice(1)}:
            <select
              value={fieldMapping[appField] || ''}
              onChange={e => setFieldMapping(f => ({ ...f, [appField]: e.target.value }))}
              style={{ marginLeft: 8 }}
            >
              <option value="">(Do not map)</option>
              {selectableColumns.map(col => (
                <option key={col.name} value={col.name}>
                  {col.displayName} ({col.name})
                </option>
              ))}
            </select>
          </label>
        ))}
      </div>
    </div>
  );
}

function SharePointSelector({ token, siteId, setSiteId, listId, setListId, columns, setColumns }) {
  const [sites, setSites] = useState([]);
  const [lists, setLists] = useState([]);

  useEffect(() => {
    if (!token) return;
    fetch("https://graph.microsoft.com/v1.0/sites?search=*", {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then((res) => res.json())
      .then((data) => setSites(data.value || []));
  }, [token]);

  useEffect(() => {
    if (!token || !siteId) return;
    fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists`, {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then((res) => res.json())
      .then((data) => setLists(data.value || []));
  }, [token, siteId]);

  useEffect(() => {
    if (!token || !siteId || !listId) {
      setColumns([]);
      return;
    }
    fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns`, {
      headers: { Authorization: `Bearer ${token}` },
    })
      .then((res) => res.json())
      .then((data) => setColumns(data.value || []));
  }, [token, siteId, listId, setColumns]);

  return (
    <div>
      <div>
        <label>
          SharePoint Site:
          <select value={siteId} onChange={e => setSiteId(e.target.value)}>
            <option value="">Select a site</option>
            {sites.map((site) => (
              <option key={site.id} value={site.id}>
                {site.displayName}
              </option>
            ))}
          </select>
        </label>
      </div>
      <div>
        <label>
          SharePoint List:
          <select value={listId} onChange={e => setListId(e.target.value)}>
            <option value="">Select a list</option>
            {lists.map((list) => (
              <option key={list.id} value={list.id}>
                {list.displayName}
              </option>
            ))}
          </select>
        </label>
      </div>
    </div>
  );
}

function CreateTicketsButton({ token, selectedEmails, emails, siteId, listId, onResult, fieldMapping }) {
  const [loading, setLoading] = useState(false);

  const handleCreateTickets = async () => {
    if (!siteId || !listId || selectedEmails.length === 0) {
      alert("Please select emails, a site, and a list.");
      return;
    }
    setLoading(true);
    let results = [];
    for (let emailId of selectedEmails) {
      const email = emails.find((e) => e.id === emailId);
      if (!email) continue;
      // Format ticket number as YYYYMMDDHHmm from receivedDateTime (local, 24-hour/military time)
      let ticketNumber = "";
      if (email.receivedDateTime) {
        const dt = new Date(email.receivedDateTime);
        console.log('TicketNumber Debug:', email.receivedDateTime, dt.toString());
        const pad = (n) => n.toString().padStart(2, '0');
        ticketNumber = `${dt.getFullYear()}${pad(dt.getMonth()+1)}${pad(dt.getDate())}${pad(dt.getHours())}${pad(dt.getMinutes())}`;
      }
      const payload = {
        fields: {
          ...(fieldMapping.subject ? { [fieldMapping.subject]: email.subject } : {}),
          ...(fieldMapping.description ? { [fieldMapping.description]: email.bodyPreview } : {}),
          ...(fieldMapping.user ? { [fieldMapping.user]: email.from?.emailAddress?.address } : {}),
          ...(fieldMapping.ticketnumber ? { [fieldMapping.ticketnumber]: ticketNumber } : {}),
        },
      };
      try {
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(payload),
          }
        );
        if (res.ok) {
          results.push({ emailId, status: "success" });
        } else {
          results.push({ emailId, status: "error", error: await res.text() });
        }
      } catch (err) {
        results.push({ emailId, status: "error", error: err.message });
      }
    }
    setLoading(false);
    onResult(results);
  };

  return (
    <div>
      <button onClick={handleCreateTickets} disabled={loading}>
        {loading ? "Creating..." : "Create Tickets"}
      </button>
    </div>
  );
}

function MainApp() {
  const [useSameAccount, setUseSameAccount] = useState(true);
  const [outlookToken, setOutlookToken] = useState(null);
  const [sharepointToken, setSharepointToken] = useState(null);
  const [outlookAccount, setOutlookAccount] = useState(null);
  const [sharepointAccount, setSharepointAccount] = useState(null);
  const [selectedEmails, setSelectedEmails] = useState([]);
  const [siteId, setSiteId] = useState("");
  const [listId, setListId] = useState("");
  const [emails, setEmails] = useState([]);
  const [results, setResults] = useState([]);
  const [columns, setColumns] = useState([]);
  const [fieldMapping, setFieldMapping] = useState({});

  // Fetch emails after Outlook login
  useEffect(() => {
    if (!outlookToken) return;
    fetch("https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20", {
      headers: { Authorization: `Bearer ${outlookToken}` },
    })
      .then((res) => res.json())
      .then((data) => setEmails(data.value || []));
  }, [outlookToken]);

  // Token to use for SharePoint
  const spToken = useSameAccount ? outlookToken : sharepointToken;

  return (
    <div style={{ maxWidth: 600, margin: "40px auto", fontFamily: "sans-serif" }}>
      <h2>Outlook to SharePoint Ticket Sync</h2>
      <DualLogin
        useSameAccount={useSameAccount}
        setUseSameAccount={setUseSameAccount}
        outlookToken={outlookToken}
        setOutlookToken={setOutlookToken}
        sharepointToken={sharepointToken}
        setSharepointToken={setSharepointToken}
        setOutlookAccount={setOutlookAccount}
        setSharepointAccount={setSharepointAccount}
      />
      {(useSameAccount ? outlookToken : (outlookToken && sharepointToken)) && (
        <>
          <SharePointSelector
            token={spToken}
            siteId={siteId}
            setSiteId={setSiteId}
            listId={listId}
            setListId={setListId}
            columns={columns}
            setColumns={setColumns}
          />
          {columns.length > 0 && (
            <FieldMapping
              columns={columns}
              fieldMapping={fieldMapping}
              setFieldMapping={setFieldMapping}
            />
          )}
          <EmailList
            token={outlookToken}
            selectedEmails={selectedEmails}
            setSelectedEmails={setSelectedEmails}
          />
          <CreateTicketsButton
            token={spToken}
            selectedEmails={selectedEmails}
            emails={emails}
            siteId={siteId}
            listId={listId}
            onResult={setResults}
            fieldMapping={fieldMapping}
          />
          <div style={{ marginTop: 20 }}>
            {results.length > 0 && (
              <div>
                <h4>Results:</h4>
                <ul>
                  {results.map((r, i) => (
                    <li key={i}>
                      Email ID: {r.emailId} — {r.status}
                      {r.error && <span style={{ color: "red" }}> ({r.error})</span>}
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <MainApp />
    </MsalProvider>
  );
}