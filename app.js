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

    // Leerzeichen kodieren
    const quartalEncoded = quartal.replace(" ", "%20");

    // Datei liegt in: Dokumente / Arbeitszeit Katharina / Katharina Arbeitszeiten_Vucinic.xlsx
    const filePath = "Arbeitszeit%20Katharina/Katharina%20Arbeitszeiten_Vucinic.xlsx";

    // FINALER API‑Pfad
    const url =
        `https://graph.microsoft.com/v1.0/me/drive/special/documents:/${filePath}:/workbook/worksheets('${quartalEncoded}')/tables('Tabelle1')/rows`;

    console.log("Speichere nach:", url);

    await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + accessToken,
            "Content-Type": "application/json"
        },
        body: JSON.stringify({ values: [[date, start, end]] })
    });

    alert(`Gespeichert in ${quartal}!`);
}
