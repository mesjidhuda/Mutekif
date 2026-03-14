const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const AdmZip = require("adm-zip");
const QRCode = require("qrcode");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static files from public directory
app.use(express.static("public"));

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir);
}

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({
    storage: storage,
    limits: { fileSize: 50 * 1024 * 1024 }
});

// Disk storage for ID Generator (needs file paths)
const diskStorage = multer.diskStorage({
    destination: uploadsDir,
    filename: (req, file, cb) => {
        cb(null, Date.now() + "-" + file.originalname);
    }
});
const uploadDisk = multer({ storage: diskStorage });

// ============================================
// HTML ROUTES
// ============================================

// Main dashboard
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "index.html"));
});

// ID Generator page
app.get("/generator", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "generator.html"));
});

// ID Validator page
app.get("/validator", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "validator.html"));
});

// ============================================
// ID GENERATOR FUNCTIONALITY
// ============================================

function formatExcelDate(serial) {
    if (!serial) return "N/A";
    if (isNaN(serial)) return String(serial).trim();
    try {
        const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
        const yyyy = date.getFullYear();
        const mm = String(date.getMonth() + 1).padStart(2, "0");
        const dd = String(date.getDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
    } catch (e) {
        return String(serial);
    }
}

app.post(
    "/api/id-generator/process",
    uploadDisk.fields([{ name: "excel" }, { name: "zip" }]),
    async (req, res) => {
        try {
            const includePhotos = req.body.includePhotos === "true";

            if (!req.files.excel) {
                return res.status(400).json({ error: "Missing Excel file." });
            }

            const workbook = xlsx.readFile(req.files.excel[0].path);
            const sheetName = workbook.SheetNames[0];
            const rawData = xlsx.utils.sheet_to_json(
                workbook.Sheets[sheetName]
            );

            let photoMap = {};
            if (includePhotos && req.files.zip && req.files.zip[0]) {
                const zip = new AdmZip(req.files.zip[0].path);
                const zipEntries = zip.getEntries();

                zipEntries.forEach(entry => {
                    if (
                        !entry.isDirectory &&
                        entry.entryName.match(/\.(jpg|jpeg|png)$/i)
                    ) {
                        const fileNameWithoutExt = entry.name
                            .split(".")
                            .slice(0, -1)
                            .join(".")
                            .toLowerCase();
                        photoMap[
                            fileNameWithoutExt
                        ] = `data:image/png;base64,${entry
                            .getData()
                            .toString("base64")}`;
                    }
                });
            }

            const processedRecords = await Promise.all(
                rawData.map(async row => {
                    const idKey = String(row.id_number || "").trim();

                    const qrPayload = JSON.stringify({
                        id: idKey,
                        name: row.full_name,
                        gender: row.gender || "N/A",
                        role: row.role || "Mutekif",
                        org: row.organization || "Mesjid Huda",
                        expiry: formatExcelDate(row.expiry_date)
                    });

                    const qrCodeBase64 = await QRCode.toDataURL(qrPayload, {
                        errorCorrectionLevel: "M",
                        margin: 1,
                        width: 300
                    });

                    return {
                        ...row,
                        expiry_date: formatExcelDate(row.expiry_date),
                        gender: row.gender || "N/A",
                        photoBase64: includePhotos
                            ? photoMap[idKey.toLowerCase()] || null
                            : null,
                        qrBase64: qrCodeBase64
                    };
                })
            );

            fs.unlinkSync(req.files.excel[0].path);
            if (includePhotos && req.files.zip && req.files.zip[0]) {
                fs.unlinkSync(req.files.zip[0].path);
            }

            res.json({ success: true, data: processedRecords });
        } catch (error) {
            console.error("ID Generator Error:", error);
            res.status(500).json({
                error: "Processing failed. Check file structures."
            });
        }
    }
);

// ============================================
// ID LOOKUP/VALIDATOR FUNCTIONALITY
// ============================================

const datasets = new Map();

setInterval(
    () => {
        const oneHourAgo = Date.now() - 60 * 60 * 1000;
        for (const [id, dataset] of datasets) {
            if (dataset.timestamp < oneHourAgo) {
                datasets.delete(id);
                console.log(`Cleaned up dataset: ${id}`);
            }
        }
    },
    60 * 60 * 1000
);

function formatFanNumber(fanNumber) {
    if (!fanNumber) return null;
    const cleaned = fanNumber.toString().replace(/\s+/g, "").replace(/\D/g, "");
    if (cleaned.length === 16) {
        return cleaned.match(/.{1,4}/g).join(" ");
    }
    return fanNumber.toString().trim();
}

// Helper function to handle duplicates and return all matches
function processRecordsWithDuplicates(
    data,
    headers,
    nameIndex,
    phoneIndex,
    fanIndex,
    idIndex
) {
    const phoneMap = new Map(); // Will store arrays of records
    const fanMap = new Map(); // Will store arrays of records
    const invalidRows = [];

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        const fullName =
            nameIndex !== -1 && row[nameIndex]
                ? String(row[nameIndex]).trim()
                : "";
        const phoneNumber =
            phoneIndex !== -1 && row[phoneIndex]
                ? String(row[phoneIndex]).trim()
                : "";
        const fanNumber =
            fanIndex !== -1 && row[fanIndex]
                ? String(row[fanIndex]).trim()
                : "";
        const idNumber =
            idIndex !== -1 && row[idIndex] ? String(row[idIndex]).trim() : "";

        if (phoneNumber && idNumber) {
            const cleanPhone = phoneNumber.replace(/\D/g, "");

            if (cleanPhone.length >= 10) {
                const formattedFan = fanNumber
                    ? formatFanNumber(fanNumber)
                    : "";

                const record = {
                    fullName: fullName || "",
                    phoneNumber: cleanPhone,
                    fanNumber: formattedFan || fanNumber || "",
                    idNumber: idNumber,
                    rowNumber: i + 1 // Track row number for reference
                };

                // Handle phone duplicates - store array
                if (!phoneMap.has(cleanPhone)) {
                    phoneMap.set(cleanPhone, []);
                }
                phoneMap.get(cleanPhone).push(record);

                // Handle Fan number duplicates - store array
                if (fanNumber && fanNumber.trim() !== "") {
                    // Store original
                    if (!fanMap.has(fanNumber)) {
                        fanMap.set(fanNumber, []);
                    }
                    fanMap.get(fanNumber).push(record);

                    // Store formatted if different
                    if (formattedFan && formattedFan !== fanNumber) {
                        if (!fanMap.has(formattedFan)) {
                            fanMap.set(formattedFan, []);
                        }
                        fanMap.get(formattedFan).push(record);
                    }

                    // Store without spaces
                    const noSpaces = fanNumber.replace(/\s+/g, "");
                    if (noSpaces !== fanNumber && noSpaces !== formattedFan) {
                        if (!fanMap.has(noSpaces)) {
                            fanMap.set(noSpaces, []);
                        }
                        fanMap.get(noSpaces).push(record);
                    }
                }
            } else {
                invalidRows.push({
                    row: i + 1,
                    phone: phoneNumber,
                    reason: "Invalid phone format"
                });
            }
        } else {
            invalidRows.push({
                row: i + 1,
                phone: phoneNumber || "missing",
                reason: !phoneNumber
                    ? "Missing phone number"
                    : "Missing ID number"
            });
        }
    }

    return {
        phoneMap,
        fanMap,
        recordCount: phoneMap.size, // Number of unique phone numbers
        totalRecords: Array.from(phoneMap.values()).reduce(
            (sum, arr) => sum + arr.length,
            0
        ), // Total records including duplicates
        invalidRows
    };
}

function parseSpreadsheet(buffer, mimetype) {
    try {
        const workbook = xlsx.read(buffer, { type: "buffer" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(firstSheet, { header: 1 });

        if (data.length < 2) {
            throw new Error(
                "Spreadsheet must contain headers and at least one row of data"
            );
        }

        const headers = data[0].map(h => String(h).toLowerCase().trim());

        const nameIndex = headers.findIndex(
            h =>
                h.includes("full_name") ||
                h.includes("fullname") ||
                h.includes("name")
        );
        const phoneIndex = headers.findIndex(h => h.includes("phone"));
        const fanIndex = headers.findIndex(
            h =>
                h.includes("fan") ||
                h.includes("national_id") ||
                h.includes("nationalid") ||
                h.includes("national")
        );
        const idIndex = headers.findIndex(
            h => h.includes("id") && !h.includes("national")
        );

        if (phoneIndex === -1) {
            throw new Error('Spreadsheet must contain a "phone_number" column');
        }
        if (idIndex === -1) {
            throw new Error('Spreadsheet must contain an "id_number" column');
        }

        // Process records with duplicate handling
        const result = processRecordsWithDuplicates(
            data,
            headers,
            nameIndex,
            phoneIndex,
            fanIndex,
            idIndex
        );

        // Calculate duplicate stats
        let phoneDuplicates = 0;
        let fanDuplicates = 0;

        for (const [key, records] of result.phoneMap) {
            if (records.length > 1) phoneDuplicates += records.length - 1;
        }

        for (const [key, records] of result.fanMap) {
            if (records.length > 1) fanDuplicates += records.length - 1;
        }

        return {
            success: true,
            phoneMap: result.phoneMap,
            fanMap: result.fanMap,
            recordCount: result.recordCount,
            totalRecords: result.totalRecords,
            phoneDuplicates,
            fanDuplicates,
            invalidRows: result.invalidRows
        };
    } catch (error) {
        console.error("Parse error:", error);
        return {
            success: false,
            error: error.message
        };
    }
}

app.post("/api/lookup/upload", upload.single("spreadsheet"), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: "No file uploaded" });
        }

        const result = parseSpreadsheet(req.file.buffer, req.file.mimetype);

        if (!result.success) {
            return res.status(400).json({ error: result.error });
        }

        const datasetId = crypto.randomBytes(16).toString("hex");

        datasets.set(datasetId, {
            phoneData: result.phoneMap,
            fanData: result.fanMap,
            timestamp: Date.now(),
            filename: req.file.originalname,
            recordCount: result.recordCount,
            totalRecords: result.totalRecords,
            phoneDuplicates: result.phoneDuplicates,
            fanDuplicates: result.fanDuplicates
        });

        res.json({
            success: true,
            datasetId,
            recordCount: result.recordCount,
            totalRecords: result.totalRecords,
            phoneDuplicates: result.phoneDuplicates,
            fanDuplicates: result.fanDuplicates,
            message: `Successfully loaded ${result.totalRecords} records (${result.recordCount} unique phone numbers)`,
            invalidRows: result.invalidRows.length
        });
    } catch (error) {
        console.error("Upload error:", error);
        res.status(500).json({ error: "Server error during upload" });
    }
});

app.post('/api/lookup/search', (req, res) => {
    try {
        const { datasetId, searchTerm, searchType } = req.body;
        
        if (!datasetId || !searchTerm || !searchType) {
            return res.status(400).json({ error: 'Missing required fields' });
        }
        
        const dataset = datasets.get(datasetId);
        if (!dataset) {
            return res.status(404).json({ error: 'Dataset not found or expired' });
        }
        
        let records = null;
        
        if (searchType === 'phone') {
            const cleanPhone = searchTerm.replace(/\D/g, '');
            records = dataset.phoneData.get(cleanPhone);
        } else {
            records = dataset.fanData.get(searchTerm);
            if (!records) {
                const noSpaces = searchTerm.replace(/\s+/g, '');
                records = dataset.fanData.get(noSpaces);
            }
            if (!records) {
                const formatted = formatFanNumber(searchTerm);
                if (formatted && formatted !== searchTerm) {
                    records = dataset.fanData.get(formatted);
                }
            }
        }
        
        if (records && records.length > 0) {
            res.json({
                found: true,
                multiple: records.length > 1,
                count: records.length,
                records: records.map(record => ({
                    fullName: record.fullName || '',
                    phoneNumber: record.phoneNumber,
                    fanNumber: record.fanNumber || '',
                    idNumber: record.idNumber,
                    rowNumber: record.rowNumber
                }))
            });
        } else {
            res.json({
                found: false,
                message: 'No record found'
            });
        }
        
    } catch (error) {
        console.error('Lookup error:', error);
        res.status(500).json({ error: 'Server error during lookup' });
    }
});

app.get("/api/lookup/dataset/:id", (req, res) => {
    const dataset = datasets.get(req.params.id);
    if (!dataset) {
        return res.status(404).json({ error: "Dataset not found" });
    }
    res.json({
        id: req.params.id,
        recordCount: dataset.recordCount,
        filename: dataset.filename,
        timestamp: dataset.timestamp
    });
});

app.get("/api/health", (req, res) => {
    res.json({
        status: "healthy",
        timestamp: new Date().toISOString(),
        activeDatasets: datasets.size
    });
});

app.use((err, req, res, next) => {
    console.error("Unhandled error:", err);
    res.status(500).json({ error: "Internal server error" });
});

app.listen(PORT, () => {
    console.log(`\n🚀 ETIKAF ADMIN SUITE - UNIFIED SYSTEM`);
    console.log(`=======================================`);
    console.log(`🌐 URL: http://localhost:${PORT}`);
    console.log(`📋 ID Generator: http://localhost:${PORT}/generator`);
    console.log(`🔍 ID Validator: http://localhost:${PORT}/validator`);
});
