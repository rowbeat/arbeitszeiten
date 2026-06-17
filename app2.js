// ---------------------------
// LOGIN (PKCE FLOW)
// ---------------------------

async function login() {
    const clientId = "3b4ebf1f-8107-477c-9368-fb76ba2169c7"; 
    const redirectUri = window.location.origin + "/arbeitszeiten/";
    const tenant = "e19f9f5c-569f-487f-8ca5-5fc648224536";

    // PKCE Code Verifier erzeugen
    const codeVerifier = [...crypto.getRandomValues(new Uint8Array(32))]
        .map(b => ("0" + b.toString(16)).slice(-2))
        .join("");

    const encoder = new TextEncoder();
    const codeChallengeBuffer = await crypto.subtle.digest("SHA-256", encoder.encode(codeVerifier));
    const codeChallenge = btoa(String.fromCharCode(...new Uint8Array(codeChallengeBuffer)))
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");

    localStorage.setItem("code_verifier", codeVerifier);

    const authUrl =
        `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize` +
        `?client_id=${clientId}` +
        `&response_type=code` +
        `&redirect_uri=${encodeURIComponent(redirectUri)}` +
        `&response_mode=query` +
        `&scope=${encodeURIComponent("openid profile offline_access Files.ReadWrite")}` +
        `&code_challenge=${codeChallenge}` +
        `&code_challenge_method=S256`;

    window.location.href = authUrl;
}

// ---------------------------
// TOKEN VERARBEITEN
// ---------------------------

async function handleRedirect() {
    const params = new URLSearchParams(window.location.search);
    const code = params.get("code");
    if (!code) return;

    const codeVerifier = localStorage.getItem("code_verifier");

    const tokenResponse = await fetch("https://login.microsoftonline.com/e19f9f5c-569f-487f-8ca5-5fc648224536/oauth2/v2.0/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
            client_id: "3b4ebf1f-8107-477c-9368-fb76ba2169c7",
            grant_type: "authorization_code",
            code: code,
            redirect_uri: window.location.origin + "/arbeitszeiten/",
            code_verifier: codeVerifier
        })
    });

    const data = await tokenResponse.json();

    if (data.access_token) {
        localStorage.setItem("access_token", data.access_token);
        alert("Login erfolgreich!");
    } else {
        alert("Login fehlgeschlagen!");
    }

    window.history.replaceState({}, document.title, "/arbeitszeiten/");
}

handleRedirect();

// ---------------------------
// SPEICHERN IN EXCEL
// ---------------------------

async function saveToExcel() {
    const accessToken = localStorage.getItem("access_token");
    if (!accessToken) {
        alert("Bitte zuerst einloggen!");
        return;
    }

    const date = document.getElementById("date").value;
    const start = document.getElementById("start").value;
    const end = document.getElementById("end").value;

    if (!date || !start || !end) {
        alert("Bitte alle Felder ausfüllen!");
        return;
    }

    // Datum → Monat → Quartal bestimmen
    const month = new Date(date).getMonth() + 1;
    let quartal = "";

    if (month >= 1 && month <= 3) quartal = "Quartal 1";
    else if (month >= 4 && month <= 6) quartal = "Quartal 2";
    else if (month >= 7 && month <= 9) quartal = "Quartal 3";
    else if (month >= 10 && month <= 12) quartal = "Quartal 4";

    const quartalEncoded = quartal.replace(" ", "%20");

    // Datei liegt in: Dokumente / Arbeitszeit Katharina / Katharina Arbeitszeiten_Vucinic.xlsx
    const filePath = "Arbeitszeit%20Katharina/Katharina%20Arbeitszeiten_Vucinic.xlsx";

    const url =
        `https://graph.microsoft.com/v1.0/me/drive/special/documents:/${filePath}:/workbook/worksheets('${quartalEncoded}')/tables('Tabelle1')/rows`;

    console.log("Speichere nach:", url);

    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + accessToken,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: [[date, start, end]] })
    });

    if (response.ok) {
        alert(`Gespeichert in ${quartal}!`);
    } else {
        const err = await response.json();
        console.error(err);
        alert("Fehler beim Speichern!");
    }
}
