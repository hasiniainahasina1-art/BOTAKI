import XLSX from 'xlsx';

export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Méthode non autorisée' });

    const { fileBase64, fileName, mode } = req.body; // mode = 'replace' ou 'append'
    if (!fileBase64 || !fileName || !mode) {
        return res.status(400).json({ error: 'Paramètres manquants' });
    }

    const token = process.env.GITHUB_TOKEN;
    if (!token) {
        return res.status(500).json({ error: 'Token GitHub non configuré' });
    }

    const repoOwner = 'hasiniainahasina1-art';
    const repoName = 'BOTAKI';
    const filePath = fileName; // database.xlsx ou CADENCIER.xlsx (à la racine)

    try {
        const newBuffer = Buffer.from(fileBase64, 'base64');

        if (mode === 'replace') {
            // Remplacement simple
            const content = newBuffer.toString('base64');
            await commitFile(token, repoOwner, repoName, filePath, content, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }
        else if (mode === 'append') {
            // 1. Récupérer le fichier existant
            const existingContent = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingContent) {
                return res.status(404).json({ error: 'Fichier existant introuvable' });
            }
            // 2. Parser les deux fichiers
            const oldWorkbook = XLSX.read(existingContent, { type: 'buffer' });
            const newWorkbook = XLSX.read(newBuffer, { type: 'buffer' });

            const oldSheet = oldWorkbook.Sheets[oldWorkbook.SheetNames[0]];
            const newSheet = newWorkbook.Sheets[newWorkbook.SheetNames[0]];

            // Convertir en tableau de lignes (en-tête compris)
            const oldData = XLSX.utils.sheet_to_json(oldSheet, { header: 1, defval: '' });
            const newData = XLSX.utils.sheet_to_json(newSheet, { header: 1, defval: '' });

            // On considère que la première ligne de l'ancien fichier est l'en-tête (si présent)
            // On garde l'en-tête de l'ancien fichier, et on ajoute les lignes du nouveau (en ignorant son éventuel en-tête)
            const hasHeader = oldData.length > 0 && oldData[0].some(cell => String(cell).toLowerCase().includes('code') || String(cell).toLowerCase().includes('barre'));
            let header = oldData[0];
            let oldBody = hasHeader ? oldData.slice(1) : oldData;

            // Pour le nouveau, on suppose la même structure (mêmes colonnes)
            let newBody = newData;
            if (hasHeader && newData.length > 0 && newData[0].some(cell => String(cell).toLowerCase().includes('code'))) {
                // Le nouveau fichier a un en-tête, on l'ignore
                newBody = newData.slice(1);
            }

            // Fusion (concaténation simple, sans dédoublonnage pour l'instant)
            const mergedBody = [...oldBody, ...newBody];

            // Reconstruire le fichier Excel
            const mergedData = [header, ...mergedBody];
            const newWs = XLSX.utils.aoa_to_sheet(mergedData);
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newWs, oldWorkbook.SheetNames[0]);

            const outBuffer = XLSX.write(newWb, { type: 'buffer', bookType: 'xlsx' });
            const content = outBuffer.toString('base64');
            await commitFile(token, repoOwner, repoName, filePath, content, `Ajout de produits dans ${fileName}`);
            return res.status(200).json({ success: true });
        }
        else {
            return res.status(400).json({ error: 'Mode invalide' });
        }
    } catch (error) {
        console.error('Erreur upload-base:', error);
        return res.status(500).json({ error: error.message });
    }
}

// Fonction pour récupérer le contenu d'un fichier depuis GitHub
async function getFileContent(token, owner, repo, path) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const resp = await fetch(url, {
        headers: { Authorization: `token ${token}` }
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    // Le contenu est encodé en base64 dans la réponse
    const content = Buffer.from(data.content, 'base64');
    return content;
}

// Fonction pour créer ou mettre à jour un fichier
async function commitFile(token, owner, repo, path, contentBase64, message) {
    // Vérifier si le fichier existe pour obtenir son sha (nécessaire pour mise à jour)
    const sha = await getFileSha(token, owner, repo, path);
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const body = {
        message: message,
        content: contentBase64,
        branch: 'main'
    };
    if (sha) {
        body.sha = sha;
    }
    const resp = await fetch(url, {
        method: 'PUT',
        headers: {
            Authorization: `token ${token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
    });
    if (!resp.ok) {
        const err = await resp.json();
        throw new Error(err.message || 'Erreur GitHub');
    }
}

async function getFileSha(token, owner, repo, path) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const resp = await fetch(url, {
        headers: { Authorization: `token ${token}` }
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    return data.sha;
}
