const express = require('express');
const archiver = require('archiver');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware pour analyser le JSON et servir les fichiers statiques
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'))); // Dossier contenant vos fichiers HTML/CSS

// ----------------------------------------------------
// 1. API - Récupération des documents des parfums
// ----------------------------------------------------
app.get('/api/perfumes-docs', (req, res) => {
    try {
        // Exemple de données (à adapter selon votre base de données réelle)
        const perfumesData = [
            {
                name: "Rose Élixir",
                files: {
                    INCI: { url: "/uploads/rose_inci.pdf", path: "./uploads/rose_inci.pdf" },
                    COA: { url: "/uploads/rose_coa.pdf", path: "./uploads/rose_coa.pdf" },
                    MSDS: null // Exemple de document manquant
                }
            },
            {
                name: "Ambre Noir",
                files: {
                    INCI: { url: "/uploads/ambre_inci.pdf", path: "./uploads/ambre_inci.pdf" },
                    COA: { url: "/uploads/ambre_coa.pdf", path: "./uploads/ambre_coa.pdf" },
                    MSDS: { url: "/uploads/ambre_msds.pdf", path: "./uploads/ambre_msds.pdf" }
                }
            }
        ];
        res.json(perfumesData);
    } catch (error) {
        res.status(500).json({ error: "Erreur lors de la récupération des données de parfums" });
    }
});

// ----------------------------------------------------
// 2. API - CRM Clients (Route qui posait problème)
// ----------------------------------------------------
app.get('/api/crm/clients', (req, res) => {
    try {
        // Remplacez ceci par votre requête de base de données (ex: MySQL, MongoDB, etc.)
        const clients = [
            { id: 1, name: "Client A", email: "clientA@example.com" },
            { id: 2, name: "Client B", email: "clientB@example.com" }
        ];
        
        // S'assure de toujours renvoyer un tableau JSON valide
        res.json(clients);
    } catch (error) {
        console.error("Erreur serveur CRM :", error);
        res.status(500).json({ error: "Erreur interne du serveur lors de la récupération des clients" });
    }
});

// ----------------------------------------------------
// 3. API - Téléchargement multiple en ZIP
// ----------------------------------------------------
app.post('/download-zip', (req, res) => {
    const { files } = req.body; // Tableau des chemins de fichiers reçus du front

    if (!files || !Array.isArray(files) || files.length === 0) {
        return res.status(400).send("Aucun fichier spécifié.");
    }

    // Configuration de l'en-tête pour envoyer un fichier ZIP
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename=Documents_Selectionnes.zip');

    const archive = archiver('zip', {
        zlib: { level: 9 } // Niveau de compression maximal
    });

    archive.on('error', (err) => {
        res.status(500).send({ error: err.message });
    });

    // Pipe l'archive directement vers la réponse HTTP
    archive.pipe(res);

    // Ajouter chaque fichier existant au ZIP
    files.forEach(filePath => {
        // Sécurisation basique des chemins pour éviter les attaques de traversée de répertoire
        const safePath = path.resolve(__dirname, filePath);
        if (fs.existsSync(safePath)) {
            archive.file(safePath, { name: path.basename(safePath) });
        }
    });

    archive.finalize();
});

// Démarrage du serveur
app.listen(PORT, () => {
    console.log(`Serveur démarré sur le port ${PORT}`);
});
