const express = require("express");
const cors = require("cors");
const axios = require("axios");
const XLSX = require("xlsx");

const app = express();
app.use(cors());
app.use(express.json());

app.get("/api/download-excel", (req, res) => {
    const url = "https://go.microsoft.com/fwlink/?LinkID=521962";

    axios.get(url, { responseType: "arraybuffer" })
        .then(response => {
            const workbook = XLSX.read(response.data, { type: "buffer" });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            let jsonData = XLSX.utils.sheet_to_json(worksheet);

            let salesKey = Object.keys(jsonData[0]).find((key) => key.trim().toLowerCase() === "sales");

            let filteredData = jsonData.filter((row) => {
                if (!row[salesKey]) return false;
                let salesValue = parseFloat(String(row[salesKey]).replace(/[$,]/g, ""));
                return !isNaN(salesValue) && salesValue > 50000;
            });

            const newWorkbook = XLSX.utils.book_new();
            const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Filtered Data");

            const excelBuffer = XLSX.write(newWorkbook, { type: "buffer", bookType: "xlsx" });

            res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            res.setHeader("Content-Disposition", 'attachment; filename="Filtered_Sales.xlsx"');
            res.setHeader("Content-Length", excelBuffer.length);

            res.send(Buffer.from(excelBuffer));
        })
        .catch(error => {
            console.error("Error downloading file:", error);
            res.status(500).send("Error fetching file");
        });
});



app.get("/api/query-base-array-input", (req, res) => {
    const url = "https://share.shub.edu.vn/api/intern-test/input";

    axios.get(url)
        .then(response => {
            res.json(response.data);
        })
        .catch(error => {
            console.error("Error when calling API:", error.message);
            res.status(500).json({ error: "Không thể lấy dữ liệu" });
        });
});


app.post("/api/query-base-array-output", async (req, res) => { 
    const { token, result } = req.body;
    const url = "https://share.shub.edu.vn/api/intern-test/output";
    
    axios.post(
        url,
        { result },
        { headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" } }
    )
    .then(response => {
        res.json(response.data);
    })
    .catch(error => {
        console.error("Error forwarding request:", error.response?.data || error.message);
        res.status(500).json({ message: "Server error", error: error.response?.data || error.message });
    }); 
});

app.listen(5000, () => {
    console.log("Backend server running on http://localhost:5000");
});
