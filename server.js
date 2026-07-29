const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const pdfParseModule = require('pdf-parse');
const archiver = require('archiver');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType } = require('docx');
const PDFDocument = require('pdfkit');

// --- PARSER PDF ---
const RawPdfParser = typeof pdfParseModule === 'function' 
    ? pdfParseModule 
    : (pdfParseModule.default || pdfParseModule.pdf || pdfParseModule);

async function parsePdfBuffer(buffer) {
    if (typeof RawPdfParser !== 'function') return null;
    try {
        return await RawPdfParser(buffer);
    } catch (err) {
        if (err.message && err.message.includes("without 'new'")) {
            const instance = new RawPdfParser(buffer);
            return instance.then ? await instance : instance;
        }
        return null;
    }
}

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// --- ROUTE POUR INJECTER UN CSS CORRECTIF (CONTRASTE MENU TÉLÉPHONE / PAYS) ---
app.get('/css/fix-dropdown.css', (req, res) => {
    res.setHeader('Content-Type', 'text/css');
    res.send(`
        /* Correctif de visibilité pour les menus déroulants de téléphone et pays */
        .iti__country-list,
        .iti__country-list * {
            color: #222222 !important;
            background-color: #ffffff !important;
        }
        .iti__country-list .iti__highlight,
        .iti__country-list .iti__country:hover {
            background-color: #e5e7eb !important;
            color: #000000 !important;
        }
        .iti__country-name, 
        .iti__dial-code {
            color: #1f2937 !important;
        }
        select option {
            background-color: #ffffff !important;
            color: #1f2937 !important;
        }
    `);
});

// --- DOSSIERS & FICHIERS ---
const TARGET_FOLDERS = [path.join(__dirname, 'Documents')];
const TEMPLATES_FOLDER = path.join(__dirname, 'Templates');
const PRICES_FOLDER = path.join(__dirname, 'Prices');
const BRAND_PROFILES_FOLDER = path.join(__dirname, 'BrandProfiles');
const CLIENTS_FILE = path.join(__dirname, 'clients.json');

app.use(express.static(__dirname));

// Création des dossiers si inexistants
[...TARGET_FOLDERS, TEMPLATES_FOLDER, PRICES_FOLDER, BRAND_PROFILES_FOLDER].forEach(folder => {
    if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
});

// Initialisation du fichier clients JSON s'il n'existe pas
if (!fs.existsSync(CLIENTS_FILE)) {
    fs.writeFileSync(CLIENTS_FILE, JSON.stringify([], null, 2));
}

// --- CONFIGURATION MULTER ---
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        if (req.params.type === 'templates') cb(null, TEMPLATES_FOLDER);
        else if (req.params.type === 'prices') cb(null, PRICES_FOLDER);
        else if (req.params.type === 'brand-profiles') cb(null, BRAND_PROFILES_FOLDER);
        else cb(null, __dirname);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        const ext = path.extname(file.originalname);
        cb(null, file.fieldname + '-' + uniqueSuffix + ext);
    }
});
const upload = multer({ storage });

// --- ROUTES PAGES HTML ---
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/report', (req, res) => res.sendFile(path.join(__dirname, 'missing-docs.html')));
app.get('/crm', (req, res) => res.sendFile(path.join(__dirname, 'crm.html')));

// --- BASE DE DONNÉES PARFUMS ---
const KNOWN_PERFUMES = [
    "LOS ANGELES", "MOSCOW", "ONYX", "MAGIC", "MALACHITE", "GRANADA", "WISTERIA", 
    "CHERRY BLOSSOM", "CAMOUFLAGE", "ELEGANT", "STYLE", "555", "18 K", "PALACE", 
    "GRAND PALAIS", "20TH ANNIVERSARY", "DIAMOND", "PEARL", "VENICE", "97 ELYSéES", 
    "GALAXY", "CLASSIC OUD", "CLASSIC ROSE", "DAMASCUS HAIRMIST", "SILK ROSE", "MARBELLA", 
    "PARIS MUSK SET", "PATCHOULI", "TOBACCO", "DAMASCUS", "GARDENIA", "TAIEF", "KALAKESS", 
    "TAJ", "PRINCE", "AL MAJD", "AL SULTAN", "GHALI", "ASHHAB", 
    "ORYX MUSK", "ORYX ROSE", "LE JARDIN", "LA FORET", "SILK", "ORYX OUD", "ORYX AMBER", 
    "LONDON", "SMOKY OUD", "SOFT OUD", "OPERA", "MOSAIC", "MOSAIC ORIENT", "CORDOBA", 
    "STAR", "MALAGA", "MAGIC HAIRMIST", "NEW YORK", "17 TH CENTURY", "16TH CENTURY", 
    "VELVET ROSE", "VELVET", "26 ELYSéES", "SHADOW", "SAPPHIRE", "AMBER", "DAMASCUS MUSK", 
    "DAMASCUS ROSE", "FINE ART", "ART DECO", "ISTANBUL", "MARRAKECH", "CHAMPS", "JOCKEY", 
    "TULIPS", "POPPIES", "BUTTERFLY", "BUBBLES", "CAPRI", "SANTORINI", "VENICE HAIR MIST", 
    "ONYX HAIR MIST", "ELEGANT HAIR MIST", "97 ELYSéES HAIRMIST", "PALACE HAIRMIST", 
    "PARIS AMAZING", "PARIS MAGICAL", "PARIS PRIVATE", "PARIS STAR", "PARIS CHIC", 
    "PARIS FEELING", "POINTS", "MAHARANI", "CANARI", "MAGIC NIGHT", "MAGIC MOON", "LAPIS", 
    "CHATEAU", "AAA", "SIDRA WOOD", "SIDRA", "FINE OUD", "SWEET OUD", "ANDALUSIAN PALACE", 
    "ANDALUSIAN GARDEN", "ORIENTAL", "NEW ZEALAND", "AUSTRALIA", "ASIA", "JAPANESE", 
    "ORYX AMBER Scented Candle", "ORYX OUD Scented Candle", "ORYX MUSK Scented Candle", 
    "ORYX ROSES SCENTED CANDLE", "ORYX MUSK REED DIFFUSER", "ORYX ROSE REED DIFFUSER", 
    "ORYX OUD REED DIFFUSER", "ORYX AMBER REED DIFFUSER", "ORYX MUSK HAIR MIST", 
    "ORYX AMBER HAIR MIST", "ORYX ROSE HAIR MIST", "ORYX OUD HAIR MIST", 
    "ORYX MUSK ROOM SPRAY", "ORYX ROSE ROOM SPRAY", "ORYX AMBER ROOM SPRAY", 
    "ORYX OUD ROOM SPRAY", "ORYX IRIS", "ORYX ELITE", "AL JASRA", "25TH ANNIVERSARY", 
    "380", "777", "SILK OUD", "OUD FOREVER", "ORYX ELITE Scented Candle", "AL THURAYA", 
    "FLOWER MUSK", "POWDERY MUSK", "JADE", "MARINE", "ORYX PATCHOULI WOOD", "KING", 
    "JOHAR", "SHARQ", "AL SOMU", "YILDIZ", "TOPKAPI", "DOLMABAHCE", "ALMAS", "ALKANZ", 
    "MAALY", "ALDAR", "ALLIGATOR", "DISCOVERY SET OPT2", "DISCOVERY SET OPT 1", "HERITAGE", 
    "SMOKY AMBER", "WHITE AMBER", "AL RAYYAN", "AL BAIRAQ", "COOKIE", "CANDY", "FIFA 2026 CANADA", 
    "FIFA 2026 USA", "FIFA 2026 MEXICO", "Ybry Turquoise", "Ybry Améthyste", "Ybry Ambre", 
    "Ybry Rubis", "Ybry Moissanite", "Ybry Émeraude", "Ybry Quartz Rose", "MAGIC PARIS", 
    "MAGIC BARCELONA", "FIFA 2026 ELITE Series", "FIFA 2026 BOX", "FIFA DISCOVERY SET", "OPAL"
];

// --- FONCTIONS UTILITAIRES DOCUMENTS ---
function getAllFilesRecursive(dirPath, arrayOfFiles = []) {
    if (!fs.existsSync(dirPath)) return arrayOfFiles;
    const files = fs.readdirSync(dirPath, { withFileTypes: true });

    files.forEach(file => {
        const fullPath = path.join(dirPath, file.name);
        if (file.isDirectory()) {
            getAllFilesRecursive(fullPath, arrayOfFiles);
        } else if (path.extname(file.name).toLowerCase() === '.pdf') {
            arrayOfFiles.push(fullPath);
        }
    });
    return arrayOfFiles;
}

async function detectDocType(filePath) {
    const fileNameUpper = path.basename(filePath).toUpperCase();

    if (fileNameUpper.includes('INCI') || fileNameUpper.includes('INGREDIENT') || fileNameUpper.includes('ALLERGEN')) return 'INCI';
    if (fileNameUpper.includes('COA') || fileNameUpper.includes('ANALYSIS') || fileNameUpper.includes('CERTIFICAT')) return 'COA';
    if (fileNameUpper.includes('MSDS') || fileNameUpper.includes('FDS') || fileNameUpper.includes('SAFETY')) return 'MSDS';

    try {
        const buffer = fs.readFileSync(filePath);
        const pdfData = await parsePdfBuffer(buffer);
        if (pdfData && pdfData.text) {
            const textUpper = pdfData.text.toUpperCase();
            if (textUpper.includes('SAFETY DATA SHEET') || textUpper.includes('FDS') || textUpper.includes('MSDS')) return 'MSDS';
            if (textUpper.includes('CERTIFICATE OF ANALYSIS') || textUpper.includes('COA')) return 'COA';
            if (textUpper.includes('INCI') || textUpper.includes('INGREDIENTS')) return 'INCI';
        }
    } catch (e) {}

    return null;
}

function matchPerfume(filePath) {
    const normalize = str => String(str).toLowerCase().replace(/[^a-z0-9]/g, '');
    const cleanPath = normalize(filePath);
    const sortedPerfumes = [...KNOWN_PERFUMES].sort((a, b) => String(b).length - String(a).length);

    for (const perfume of sortedPerfumes) {
        if (cleanPath.includes(normalize(perfume))) return perfume;
    }
    return null;
}

async function generateReportData() {
    let allPdfPaths = [];
    TARGET_FOLDERS.forEach(folder => {
        allPdfPaths = allPdfPaths.concat(getAllFilesRecursive(folder));
    });

    const perfumesMap = {};
    KNOWN_PERFUMES.forEach(p => {
        const displayName = String(p).charAt(0).toUpperCase() + String(p).slice(1).toLowerCase();
        perfumesMap[p] = { name: displayName, files: {} };
    });

    for (const filePath of allPdfPaths) {
        const matchedPerfume = matchPerfume(filePath);
        if (matchedPerfume) {
            const docType = await detectDocType(filePath);
            if (docType && !perfumesMap[matchedPerfume].files[docType]) {
                perfumesMap[matchedPerfume].files[docType] = true;
            }
        }
    }

    const requiredDocs = ['INCI', 'COA', 'MSDS'];
    const report = [];
    const stats = { missingINCI: 0, missingCOA: 0, missingMSDS: 0 };

    Object.values(perfumesMap).forEach(item => {
        const missing = [];
        const present = [];

        requiredDocs.forEach(doc => {
            if (!item.files[doc]) {
                missing.push(doc);
                stats[`missing${doc}`]++;
            } else {
                present.push(doc);
            }
        });

        if (missing.length > 0) {
            report.push({ name: item.name, missing, present });
        }
    });

    return { incompleteCount: report.length, stats, report };
}

// --- FONCTIONS UTILITAIRES CRM ---
function getClients() {
    try {
        if (!fs.existsSync(CLIENTS_FILE)) return [];
        const data = fs.readFileSync(CLIENTS_FILE, 'utf8');
        return JSON.parse(data);
    } catch (e) {
        console.error("Erreur lecture clients.json:", e);
        return [];
    }
}

function saveClients(clients) {
    try {
        fs.writeFileSync(CLIENTS_FILE, JSON.stringify(clients, null, 2), 'utf8');
        return true;
    } catch (e) {
        console.error("Erreur écriture clients.json:", e);
        return false;
    }
}

// --- ENDPOINTS FICHIERS & RAPPORTS ---
app.post('/api/upload/:type', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier fourni' });
    res.json({ success: true, fileName: req.file.filename });
});

app.get('/api/sidebar-files', (req, res) => {
    const getFolderFiles = (folderPath) => {
        if (!fs.existsSync(folderPath)) return [];
        return fs.readdirSync(folderPath).map(file => ({
            name: file,
            url: `/download?path=${encodeURIComponent(path.join(folderPath, file))}`
        }));
    };

    res.json({
        templates: getFolderFiles(TEMPLATES_FOLDER),
        prices: getFolderFiles(PRICES_FOLDER)
    });
});

app.get('/api/perfumes-docs', async (req, res) => {
    try {
        let allPdfPaths = [];
        TARGET_FOLDERS.forEach(folder => {
            allPdfPaths = allPdfPaths.concat(getAllFilesRecursive(folder));
        });

        const perfumesMap = {};
        KNOWN_PERFUMES.forEach(p => {
            const displayName = String(p).charAt(0).toUpperCase() + String(p).slice(1).toLowerCase();
            perfumesMap[p] = { name: displayName, files: {} };
        });

        for (const filePath of allPdfPaths) {
            const matchedPerfume = matchPerfume(filePath);
            if (matchedPerfume) {
                const docType = await detectDocType(filePath);
                if (docType && !perfumesMap[matchedPerfume].files[docType]) {
                    perfumesMap[matchedPerfume].files[docType] = {
                        url: `/download?path=${encodeURIComponent(filePath)}`,
                        path: filePath
                    };
                }
            }
        }

        const result = Object.values(perfumesMap).filter(item => Object.keys(item.files).length > 0);
        res.json(result);
    } catch (err) {
        res.status(500).json({ error: "Erreur lors du scan" });
    }
});

app.get('/api/missing-docs-report', async (req, res) => {
    try {
        const data = await generateReportData();
        res.json(data);
    } catch (err) {
        res.status(500).json({ error: "Erreur serveur" });
    }
});

// --- EXPORTS (Excel, Word, PDF) ---
app.get('/api/export/excel', async (req, res) => {
    try {
        const data = await generateReportData();
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Missing Docs');

        sheet.columns = [
            { header: 'Perfume Name', key: 'name', width: 30 },
            { header: 'Missing Documents', key: 'missing', width: 25 },
            { header: 'Present Documents', key: 'present', width: 25 },
            { header: 'Missing Count', key: 'count', width: 15 }
        ];

        data.report.forEach(item => {
            sheet.addRow({
                name: item.name,
                missing: item.missing.join(', '),
                present: item.present.join(', '),
                count: `${item.missing.length} / 3`
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="Missing_Documents_Report.xlsx"');
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        res.status(500).send("Erreur génération Excel");
    }
});

app.get('/api/export/word', async (req, res) => {
    try {
        const data = await generateReportData();
        const tableRows = [
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ text: "Perfume Name", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Missing Documents", bold: true })] }),
                    new TableCell({ children: [new Paragraph({ text: "Present Documents", bold: true })] }),
                ]
            })
        ];

        data.report.forEach(item => {
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(item.name)] }),
                    new TableCell({ children: [new Paragraph(item.missing.join(', '))] }),
                    new TableCell({ children: [new Paragraph(item.present.join(', '))] }),
                ]
            }));
        });

        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: "Missing Documents Report", heading: "Heading1" }),
                    new Paragraph({ text: `Total Incomplete: ${data.incompleteCount} perfumes\n\n` }),
                    new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="Missing_Documents_Report.docx"');
        res.send(buffer);
    } catch (err) {
        res.status(500).send("Erreur génération Word");
    }
});

app.get('/api/export/pdf', async (req, res) => {
    try {
        const data = await generateReportData();
        const doc = new PDFDocument({ margin: 30 });

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename="Missing_Documents_Report.pdf"');

        doc.pipe(res);
        doc.fontSize(18).text('Missing Documents Report', { align: 'center' });
        doc.moveDown();
        doc.fontSize(12).text(`Total Incomplete Perfumes: ${data.incompleteCount}`);
        doc.moveDown();

        data.report.forEach(item => {
            doc.fontSize(11).text(`${item.name}`, { bold: true });
            doc.fontSize(10).text(`  - Missing: ${item.missing.join(', ')}`);
            doc.fontSize(10).text(`  - Available: ${item.present.join(', ')}`);
            doc.moveDown(0.5);
        });

        doc.end();
    } catch (err) {
        res.status(500).send("Erreur génération PDF");
    }
});

// --- TÉLÉCHARGEMENTS ---
app.get('/download', (req, res) => {
    if (req.query.path && fs.existsSync(req.query.path)) res.download(req.query.path);
    else res.status(404).send('Fichier introuvable');
});

app.post('/download-zip', (req, res) => {
    try {
        const { files } = req.body;
        if (!files || !Array.isArray(files) || files.length === 0) return res.status(400).json({ error: 'Aucun fichier' });

        const validFiles = files.filter(f => f && fs.existsSync(f));
        if (validFiles.length === 0) return res.status(400).json({ error: 'Introuvable' });

        res.attachment('Documents_Parfums.zip');
        const archive = archiver('zip', { zlib: { level: 9 } });

        archive.on('error', err => res.status(500).send("Erreur ZIP"));
        archive.pipe(res);

        validFiles.forEach(filePath => {
            archive.file(filePath, { name: path.basename(filePath) });
        });

        archive.finalize();
    } catch (err) {
        res.status(500).send("Erreur serveur");
    }
});

// --- API CRM ---

// 1. Récupérer tous les clients
app.get('/api/crm/clients', (req, res) => {
    res.json(getClients());
});

// 2. Ajouter un client (multer pour supporter le FormData envoyé par le formulaire)
app.post('/api/crm/clients', upload.single('brandProfile'), (req, res) => {
    try {
        const clients = getClients();
        const brandProfile = req.file ? {
            filename: req.file.filename,
            originalName: req.file.originalname,
            url: `/download?path=${encodeURIComponent(req.file.path)}`
        } : null;
        const newClient = {
            id: Date.now().toString(),
            name: req.body.name || 'Sans Nom',
            company: req.body.company || '',
            email: req.body.email || '',
            phone: req.body.phone || '',
            phone2: req.body.phone2 || '',
            country: req.body.country || '',
            businessType: req.body.businessType || '',
            status: req.body.status || 'Prospect',
            brandProfile,
            notes: Array.isArray(req.body.notes) ? req.body.notes : [],
            createdAt: new Date().toISOString()
        };
        clients.push(newClient);
        if (saveClients(clients)) {
            res.json({ success: true, client: newClient });
        } else {
            res.status(500).json({ error: "Erreur écriture fichier" });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 2.5. Ajouter une note à un client spécifique
app.post('/api/crm/clients/:id/notes', (req, res) => {
    try {
        let clients = getClients();
        const client = clients.find(c => c.id === req.params.id);
        if (!client) return res.status(404).json({ error: "Client introuvable" });

        if (!Array.isArray(client.notes)) {
            client.notes = [];
        }

        const newNote = {
            id: Date.now().toString(),
            text: req.body.text || req.body.note || '',
            createdAt: new Date().toISOString()
        };

        client.notes.push(newNote);
        if (saveClients(clients)) {
            res.json({ success: true, client, note: newNote });
        } else {
            res.status(500).json({ error: "Erreur écriture fichier" });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 3. Mettre à jour un client (multer pour parser le FormData + gestion brandProfile)
app.put('/api/crm/clients/:id', upload.single('brandProfile'), (req, res) => {
    try {
        let clients = getClients();
        const index = clients.findIndex(c => c.id === req.params.id);
        if (index === -1) return res.status(404).json({ error: "Client introuvable" });

        let brandProfile = clients[index].brandProfile;
        if (req.file) {
            brandProfile = {
                filename: req.file.filename,
                originalName: req.file.originalname,
                url: `/download?path=${encodeURIComponent(req.file.path)}`
            };
        }

        clients[index] = {
            ...clients[index],
            name: req.body.name !== undefined ? req.body.name : clients[index].name,
            company: req.body.company !== undefined ? req.body.company : clients[index].company,
            email: req.body.email !== undefined ? req.body.email : clients[index].email,
            phone: req.body.phone !== undefined ? req.body.phone : clients[index].phone,
            phone2: req.body.phone2 !== undefined ? req.body.phone2 : clients[index].phone2,
            country: req.body.country !== undefined ? req.body.country : clients[index].country,
            businessType: req.body.businessType !== undefined ? req.body.businessType : clients[index].businessType,
            status: req.body.status !== undefined ? req.body.status : clients[index].status,
            notes: req.body.notes !== undefined ? req.body.notes : clients[index].notes,
            brandProfile
        };

        if (saveClients(clients)) {
            res.json({ success: true, client: clients[index] });
        } else {
            res.status(500).json({ error: "Erreur d'enregistrement dans le fichier" });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 4. Supprimer un client
app.delete('/api/crm/clients/:id', (req, res) => {
    try {
        let clients = getClients();
        clients = clients.filter(c => c.id !== req.params.id);
        if (saveClients(clients)) {
            res.json({ success: true });
        } else {
            res.status(500).json({ error: "Erreur écriture fichier" });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5. Upload Brand Profile pour un client spécifique
app.post('/api/crm/upload-brand-profile/:clientId', upload.single('brandProfile'), (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: "Aucun fichier fourni" });

        let clients = getClients();
        const client = clients.find(c => c.id === req.params.clientId);
        if (!client) return res.status(404).json({ error: "Client non trouvé" });

        const fileData = {
            filename: req.file.filename,
            originalName: req.file.originalname,
            url: `/download?path=${encodeURIComponent(req.file.path)}`
        };

        client.brandProfile = fileData;
        if (saveClients(clients)) {
            res.json({ success: true, brandProfile: fileData });
        } else {
            res.status(500).json({ error: "Erreur écriture fichier" });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 6. Importation massive de clients
app.post('/api/crm/import-clients', (req, res) => {
    try {
        const { clientsList } = req.body;
        if (!Array.isArray(clientsList) || clientsList.length === 0) {
            return res.status(400).json({ error: "Format invalide ou tableau vide" });
        }

        const currentClients = getClients();
        let importedCount = 0;

        clientsList.forEach(item => {
            if (item.name || item.company) {
                currentClients.push({
                    id: (Date.now() + Math.random()).toString(),
                    name: item.name || item.company || 'Inconnu',
                    company: item.company || '',
                    email: item.email || '',
                    phone: item.phone || '',
                    phone2: item.phone2 || '',
                    country: item.country || '',
                    businessType: item.businessType || '',
                    status: item.status || 'Prospect',
                    brandProfile: null,
                    notes: [],
                    createdAt: new Date().toISOString()
                });
                importedCount++;
            }
        });

        if (saveClients(currentClients)) {
            res.json({ success: true, count: importedCount });
        } else {
            res.status(500).json({ error: "Erreur écriture fichier" });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// --- DÉMARRAGE DU SERVEUR (Compatible Render & Local) ---
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Anas Workbook prêt sur le port ${PORT}`);
});
