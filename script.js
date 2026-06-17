console.log("script.js wurde geladen");

// ---------------- PKCE HELFERFUNKTIONEN ----------------

function base64urlencode(str) {
    return btoa(String.fromCharCode.apply(null, new Uint8Array(str)))
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");
}

async function sha256(plain) {
    const encoder = new TextEncoder();
    const data = encoder.encode(plain);
    return await crypto.subtle.digest("SHA-256", data);
}

function generateRandomString(length) {
    const charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    let result = "";
    const values = crypto.getRandomValues(new Uint8Array(length));
    for (let i = 0; i < length; i++) result += charset[values[i] % charset.length];
    return result;
}

async function createPKCE() {
    const codeVerifier = generateRandomString(64);
    const hashed = await sha256(codeVerifier);
    const codeChallenge = base64urlencode(hashed);

    sessionStorage.setItem("code_verifier", codeVerifier);
    return codeChallenge;
}

// ---------------- LOGIN STARTEN ----------------

async function login() {
    const clientId = "3b4ebf1f-8107-477c-9368-fb76ba2169c7";
    const redirectUri = "https://rowbeat.github.io/arbeitszeiten/";
    const scope = "offline_access Files.ReadWrite Files.ReadWrite.All User.Read";

    const codeChallenge = await createPKCE();

    const authUrl =
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize" +
        "?client_id=" + clientId +
        "&response_type=code" +
        "&redirect_uri=" + encodeURIComponent(redirectUri) +
        "&scope=" + encodeURIComponent(scope) +
        "&code_challenge=" + codeChallenge +
        "&code_challenge_method=S256";

    window.location.href = authUrl;
}

// ---------------- TOKEN AUSTAUSCH ----------------

async function exchangeCodeForToken() {
    const params = new URLSearchParams(window.location.search);
    const code = params.get("code");
    if (!code) return;

    const clientId = "3b4ebf1f-8107-477c-9368-fb76ba2169c7";
    const redirectUri = "https://rowbeat.github.io/arbeitszeiten/";
    const scope = "offline_access Files.ReadWrite Files.ReadWrite.All User.Read";

    const codeVerifier = sessionStorage.getItem("code_verifier");

    const body = new URLSearchParams({
        client_id: clientId,
        grant_type: "authorization_code",
        code: code,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier,
        scope: scope
    });

    const response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body
    });

    const data = await response.json();
    console.log("Token Response:", data);

    localStorage.setItem("access_token", data.access_token);
    localStorage.setItem("refresh_token", data.refresh_token);

    // URL aufräumen
    window.history.replaceState({}, document.title, "/arbeitszeiten/");
}

exchangeCodeForToken();

// ---------------- EXCEL SPEICHERN ----------------

async function saveToExcel() {
    const accessToken = localStorage.getItem("access_token");
    if (!accessToken) {
        alert("Bitte zuerst einloggen!");
        return;
    }

    const date = document.getElementById("date").value;
    const start = document.getElementById("start").value;
    const end = document.getElementById("end").value;

    await fetch("https://graph.microsoft.com/v1.0/me/drive/root:/Arbeitszeiten_Vucinic.xlsx:/workbook/worksheets('Tabelle1')/tables('Tabelle1')/rows", {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + accessToken,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: [ [date, start, end] ] })
    });

    alert("Gespeichert!");
}
