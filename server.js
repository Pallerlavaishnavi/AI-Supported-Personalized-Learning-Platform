const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const XLSX = require("xlsx");
const fs = require("fs");

const app = express();
app.use(cors());
app.use(bodyParser.json());

const FILE = "data.xlsx";

// ---- REGISTER STUDENT ----
app.post("/register", (req, res) => {
    let data = [];

    if (fs.existsSync(FILE)) {
        const workbook = XLSX.readFile(FILE);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        data = XLSX.utils.sheet_to_json(sheet);
    }

    const { username, password, email, weakness, age, phone, studying } = req.body;

    data.push({
        username,
        password,
        email,
        weakness,
        age,
        phone,
        studying,
        video_link: "",
        pdf_link: "",
        completed: "No",
        test_score: "",
        improvement: "",
        feedback: "",
        test_status: "Not attempted"   // NEW
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Students");
    XLSX.writeFile(workbook, FILE);

    res.json({ success: true, msg: "Student registered successfully" });
});

// ---- GET STUDENTS FOR ADMIN / STUDENT PAGE ----
app.get("/students", (req, res) => {
    if (!fs.existsSync(FILE)) return res.json([]);

    try {
        const workbook = XLSX.readFile(FILE);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);
        res.json(data);
    } catch (err) {
        console.error("XLSX read error:", err);
        res.json([]);
    }
});

// ---- STUDENT LOGIN ----
app.post("/studentLogin", (req, res) => {
    const { username, password } = req.body;

    if (!fs.existsSync(FILE)) {
        return res.json({ success: false, msg: "No data file" });
    }

    const workbook = XLSX.readFile(FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    const s = data.find(st => st.username === username && st.password === password);
    if (!s) {
        return res.json({ success: false, msg: "Invalid username or password" });
    }
    res.json({ success: true, student: s });
});

// ---- ASSIGN / REASSIGN LEARNING PATH ----
app.post("/assign", (req, res) => {
    const { username, video_link, pdf_link } = req.body;

    if (!fs.existsSync(FILE)) {
        return res.json({ success: false, msg: "No student data found" });
    }

    const workbook = XLSX.readFile(FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet);

    let student = data.find(s => s.username === username);
    if (!student) {
        return res.json({ success: false, msg: "Student not found" });
    }

    student.video_link = video_link || student.video_link || "";
    student.pdf_link = pdf_link || student.pdf_link || "";

    const newSheet = XLSX.utils.json_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Students");
    XLSX.writeFile(newWorkbook, FILE);

    res.json({ success: true, msg: "Learning path assigned successfully" });
});

// ---- MARK PATH COMPLETED ----
app.post("/completePath", (req, res) => {
    const { username } = req.body;

    if (!fs.existsSync(FILE)) {
        return res.json({ success: false, msg: "No data file" });
    }

    const workbook = XLSX.readFile(FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet);

    const s = data.find(st => st.username === username);
    if (!s) return res.json({ success: false, msg: "Student not found" });

    s.completed = "Yes";

    const newSheet = XLSX.utils.json_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Students");
    XLSX.writeFile(newWorkbook, FILE);

    res.json({ success: true, msg: "Learning path marked as completed" });
});

// ---- MARK TEST SKIPPED ----
app.post("/skipTest", (req, res) => {
    const { username } = req.body;

    if (!fs.existsSync(FILE)) {
        return res.json({ success: false, msg: "No data file" });
    }

    const workbook = XLSX.readFile(FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet);

    const s = data.find(st => st.username === username);
    if (!s) return res.json({ success: false, msg: "Student not found" });

    s.test_status = "Skipped";

    const newSheet = XLSX.utils.json_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Students");
    XLSX.writeFile(newWorkbook, FILE);

    res.json({ success: true, msg: "Test marked as skipped" });
});

// ---- SUBMIT TEST & FEEDBACK ----
app.post("/submitTest", (req, res) => {
    const { username, score, improvement, feedback } = req.body;

    if (!fs.existsSync(FILE)) {
        return res.json({ success: false, msg: "No data file" });
    }

    const workbook = XLSX.readFile(FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet);

    const s = data.find(st => st.username === username);
    if (!s) return res.json({ success: false, msg: "Student not found" });

    s.test_score = score;
    s.improvement = improvement;
    s.feedback = feedback;
    s.test_status = "Completed";   // NEW

    const newSheet = XLSX.utils.json_to_sheet(data);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Students");
    XLSX.writeFile(newWorkbook, FILE);

    res.json({ success: true, msg: "Test & feedback saved" });
});

// ---- DELETE STUDENT ----
app.post("/delete", (req, res) => {
    const { username } = req.body;

    if (!fs.existsSync(FILE)) {
        return res.json({ success: false, msg: "No data file found" });
    }

    const workbook = XLSX.readFile(FILE);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet);

    const newData = data.filter(s => s.username !== username);

    const newSheet = XLSX.utils.json_to_sheet(newData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Students");
    XLSX.writeFile(newWorkbook, FILE);

    res.json({ success: true, msg: "Student deleted successfully" });
});

app.listen(3000, () => {
    console.log("🚀 Server running at http://127.0.0.1:3000");
});
