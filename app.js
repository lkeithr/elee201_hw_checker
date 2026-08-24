let workbook = null;
let solutions = [];

// Load default workbook
window.onload = function () {
    loadDefaultWorkbook();
    document.getElementById('fileInput').addEventListener('change', handleFile);
};

// Load Excel from same folder
function loadDefaultWorkbook() {
    console.log("Attempting to load workbook...");

    fetch("elee201_hw_answers_complete.xlsx")
        .then(response => {
            console.log("Fetch response:", response);
            return response.arrayBuffer();
        })
        .then(data => {
            console.log("Workbook data loaded");

            workbook = XLSX.read(data, { type: "array" });

            console.log("Sheet names:", workbook.SheetNames);

            populateDropdown();
        })
        .catch(err => {
            console.error("Error loading file:", err);
        });
}

// Populate dropdown
function populateDropdown() {
    const select = document.getElementById("sheetSelect");
    select.innerHTML = "";

    workbook.SheetNames.forEach(name => {
        const option = document.createElement("option");
        option.value = name;
        option.text = name;
        select.appendChild(option);
    });

    if (workbook.SheetNames.length > 0) {
        select.value = workbook.SheetNames[0];
        loadSheet();
    }
}

// Load selected sheet
function loadSheet() {
    const sheetName = document.getElementById("sheetSelect").value;
    if (!sheetName) return;

    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    parseSheet(data);
    buildUI();

    document.getElementById("scoreDisplay").innerText = "Score: --";
}

// Handle uploaded file
function handleFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function (evt) {
        const data = new Uint8Array(evt.target.result);
        workbook = XLSX.read(data, { type: "array" });
        populateDropdown();
    };

    reader.readAsArrayBuffer(file);
}

// Parse sheet
function parseSheet(data) {
    solutions = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];

        if (!row[0]) continue;

        if (row[0].toString().toLowerCase().includes("problem")) continue;

        const problem = row[0];
        let answers = [];

        for (let j = 1; j < row.length; j += 4) {
            if (!row[j]) continue;

            answers.push({
                name: row[j],
                value: parseFloat(row[j + 1]),
                units: row[j + 2],
                tol: parseFloat(row[j + 3])
            });
        }

        if (answers.length > 0) {
            solutions.push({ problem, answers });
        }
    }
}

// Build UI
function buildUI() {
    const container = document.getElementById("problems");
    container.innerHTML = "";

    solutions.forEach((p, pi) => {
        const div = document.createElement("div");
        div.className = "problem";

        const title = document.createElement("h3");
        title.innerText = "Problem " + p.problem;
        div.appendChild(title);

        p.answers.forEach((a, ai) => {

            const label = document.createElement("label");
            label.innerText = `${a.name} (${a.units}): `;

            const input = document.createElement("input");
            input.type = "number";
            input.step = "any";
            input.id = `p${pi}_a${ai}`;

            const result = document.createElement("span");
            result.id = `r${pi}_a${ai}`;

            div.appendChild(label);
            div.appendChild(input);
            div.appendChild(result);
            div.appendChild(document.createElement("br"));
        });

        container.appendChild(div);
    });
}

// Check answers + scoring
function checkAnswers() {

    let totalScore = 0;
    let totalProblems = solutions.length;

    solutions.forEach((p, pi) => {

        let problemPoints = 1 / totalProblems;
        let perAnswerPoints = problemPoints / p.answers.length;
        let problemScore = 0;

        p.answers.forEach((a, ai) => {

            const input = parseFloat(document.getElementById(`p${pi}_a${ai}`).value);
            const result = document.getElementById(`r${pi}_a${ai}`);

            if (isNaN(input)) {
                result.innerText = " Enter value";
                result.className = "incorrect";
                return;
            }

            if (Math.abs(input - a.value) <= a.tol) {
                result.innerText = " Correct";
                result.className = "correct";
                problemScore += perAnswerPoints;
            } else {
                result.innerText = " Incorrect";
                result.className = "incorrect";
            }

        });

        totalScore += problemScore;
    });

    let percent = (totalScore * 100).toFixed(1);

    document.getElementById("scoreDisplay").innerText =
        `Score: ${percent}%`;
}
