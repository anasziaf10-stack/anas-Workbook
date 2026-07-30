const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const pdfParseModule = require('pdf-parse');
const archiver = require('archiver');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType } = require('docx');
const PDFDocument = require('pdfkit');
const { createClient } = require('@supabase/supabase-js');

// --- CONFIGURATION SUPABASE ---
const supabaseUrl = 'https://ucyakopdwvinspmflfns.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVjeWFrb3Bkd3ZpbnNwbmZsZm5zIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4NTMzNTgyMywiZXhwIjoyMTAwOTExODIzfQ.gFX8320QdN9t7XTXBa5z7_I1MyuNTzS0BMWNFCrjO1I';
const supabase = createClient(supabaseUrl, supabaseKey);

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

// --- ROUTE DE DIAGNOSTIC RÉSEAU (temporaire) ---
app.get('/api/debug-network', async (req, res) => {
    const results = {
        nodeVersion: process.version,
        hasGlobalFetch: typeof fetch !== 'undefined',
        tests: {}
    };

    try {
        const r = await fetch(supabaseUrl + '/rest/v1/', {
            headers: { apikey: supabaseKey }
        });
        results.tests.rawFetchToSupabase = { status: r.status, ok: r.ok };
    } catch (err) {
        results.tests.rawFetchToSupabase = {
            error: err.message,
            cause: err.cause ? String(err.cause) : null,
            causeCode: err.cause && err.cause.code ? err.cause.code : null
        };
    }

    try {
        const r2 = await fetch('https://www.google.com');
        results.tests.rawFetchToGoogle = { status: r2.status, ok: r2.ok };
    } catch (err) {
        results.tests.rawFetchToGoogle = {
            error: err.message,
            cause: err.cause ? String(err.cause) : null,
            causeCode: err.cause && err.cause.code ? err.cause.code : null
        };
    }

    try {
        const dns = require('dns').promises;
        const host = new URL(supabaseUrl).hostname;
        const addresses = await dns.lookup(host, { all: true });
        results.tests.dnsLookup = addresses;
    } catch (err) {
        results.tests.dnsLookup = { error: err.message, code: err.code };
    }

    res.json(results);
});

// --- ROUTE POUR INJECTER UN CSS CORRECTIF (CONTRASTE MENU TÉLÉPHONE / PAYS) ---
app.get('/css/fix-dropdown.css', (req, res) => {
    res.setHeader('Content-Type', 'text/css');
    res.send(`
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

// --- DOSSIERS ---
const TARGET_FOLDERS = [path.join(__dirname, 'Documents')];
const TEMPLATES_FOLDER = path.join(__dirname, 'Templates');
const PRICES_FOLDER = path.join(__dirname, 'Prices');
const BRAND_PROFILES_FOLDER = path.join(__dirname, 'BrandProfiles');

app.use(express.static(__dirname));

[...TARGET_FOLDERS, TEMPLATES_FOLDER, PRICES_FOLDER, BRAND_PROFILES_FOLDER].forEach(folder => {
    if (!fs.existsSync(folder)) fs.mkdirSync(folder, { recursive: true });
});

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
        if (missing.length > 0) report.push({ name: item.name, missing, present });
    });
    return { incompleteCount: report.length, stats, report };
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
    res.json({ templates: getFolderFiles(TEMPLATES_FOLDER), prices: getFolderFiles(PRICES_FOLDER) });
});

app.get('/api/perfumes-docs', async (req, res) => {
    try {
        let allPdfPaths = [];
        TARGET_FOLDERS.forEach(folder => allPdfPaths = allPdfPaths.concat(getAllFilesRecursive(folder)));
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
                    perfumesMap[matchedPerfume].files[docType] = { url: `/download?path=${encodeURIComponent(filePath)}`, path: filePath };
                }
            }
        }
        res.json(Object.values(perfumesMap).filter(item => Object.keys(item.files).length > 0));
    } catch (err) {
        res.status(500).json({ error: "Erreur lors du scan" });
    }
});

app.get('/api/missing-docs-report', async (req, res) => {
    try {
        res.json(await generateReportData());
    } catch (err) {
        res.status(500).json({ error: "Erreur serveur" });
    }
});

// --- API CRM SUPABASE ---

// 1. Récupérer tous les clients
app.get('/api/crm/clients', async (req, res) => {
    try {
        const { data, error } = await supabase.from('clients').select('*');
        if (error) throw error;
        const formatted = data.map(c => ({
            ...c,
            businessType: c.business_type,
            brandProfile: c.brand_profile,
            createdAt: c.created_at
        }));
        res.json(formatted);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 2. Ajouter un client
app.post('/api/crm/clients', upload.single('brandProfile'), async (req, res) => {
    try {
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
            business_type: req.body.businessType || req.body.business_type || '',
            status: req.body.status || 'Prospect',
            brand_profile: brandProfile,
            notes: Array.isArray(req.body.notes) ? req.body.notes : [],
            created_at: new Date().toISOString()
        };

        const { data, error } = await supabase.from('clients').insert([newClient]).select();
        if (error) throw error;

        res.json({ success: true, client: { ...newClient, businessType: newClient.business_type, brandProfile, createdAt: newClient.created_at } });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 2.5. Ajouter une note à un client
app.post('/api/crm/clients/:id/notes', async (req, res) => {
    try {
        const { data: clientData, error: fetchError } = await supabase.from('clients').select('*').eq('id', req.params.id).single();
        if (fetchError || !clientData) return res.status(404).json({ error: "Client introuvable" });

        const notes = Array.isArray(clientData.notes) ? clientData.notes : [];
        const newNote = {
            id: Date.now().toString(),
            text: req.body.text || req.body.note || '',
            createdAt: new Date().toISOString()
        };
        notes.push(newNote);

        const { error: updateError } = await supabase.from('clients').update({ notes }).eq('id', req.params.id);
        if (updateError) throw updateError;

        res.json({ success: true, client: { ...clientData, notes }, note: newNote });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 3. Mettre à jour un client
app.put('/api/crm/clients/:id', upload.single('brandProfile'), async (req, res) => {
    try {
        const { data: existing, error: fetchErr } = await supabase.from('clients').select('*').eq('id', req.params.id).single();
        if (fetchErr || !existing) return res.status(404).json({ error: "Client introuvable" });

        let brandProfile = existing.brand_profile;
        if (req.file) {
            brandProfile = {
                filename: req.file.filename,
                originalName: req.file.originalname,
                url: `/download?path=${encodeURIComponent(req.file.path)}`
            };
        }

        const updates = {
            name: req.body.name !== undefined ? req.body.name : existing.name,
            company: req.body.company !== undefined ? req.body.company : existing.company,
            email: req.body.email !== undefined ? req.body.email : existing.email,
            phone: req.body.phone !== undefined ? req.body.phone : existing.phone,
            phone2: req.body.phone2 !== undefined ? req.body.phone2 : existing.phone2,
            country: req.body.country !== undefined ? req.body.country : existing.country,
            business_type: req.body.businessType !== undefined ? req.body.businessType : (req.body.business_type !== undefined ? req.body.business_type : existing.business_type),
            status: req.body.status !== undefined ? req.body.status : existing.status,
            notes: req.body.notes !== undefined ? req.body.notes : existing.notes,
            brand_profile: brandProfile
        };

        const { data, error } = await supabase.from('clients').update(updates).eq('id', req.params.id).select().single();
        if (error) throw error;

        res.json({ success: true, client: { ...data, businessType: data.business_type, brandProfile: data.brand_profile, createdAt: data.created_at } });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 4. Supprimer un client
app.delete('/api/crm/clients/:id', async (req, res) => {
    try {
        const { error } = await supabase.from('clients').delete().eq('id', req.params.id);
        if (error) throw error;
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5. Importation massive
app.post('/api/crm/import-clients', async (req, res) => {
    try {
        const { clientsList } = req.body;
        if (!Array.isArray(clientsList) || clientsList.length === 0) {
            return res.status(400).json({ error: "Format invalide ou tableau vide" });
        }

        const rowsToInsert = clientsList.filter(item => item.name || item.company).map(item => ({
            id: (Date.now() + Math.random()).toString(),
            name: item.name || item.company || 'Inconnu',
            company: item.company || '',
            email: item.email || '',
            phone: item.phone || '',
            phone2: item.phone2 || '',
            country: item.country || '',
            business_type: item.businessType || item.business_type || '',
            status: item.status || 'Prospect',
            brand_profile: null,
            notes: [],
            created_at: new Date().toISOString()
        }));

        const { error } = await supabase.from('clients').insert(rowsToInsert);
        if (error) throw error;

        res.json({ success: true, count: rowsToInsert.length });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// --- DÉMARRAGE DU SERVEUR ---
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`Anas Workbook avec Supabase prêt sur le port ${PORT}`);
});
