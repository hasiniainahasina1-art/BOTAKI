import ExcelJS from 'exceljs';

// Définition des colonnes attendues pour chaque type de fichier
const FILE_STRUCTURE = {
    'database.xlsx': ['Code barre', 'Code article', 'Désignation'],
    'CADENCIER.xlsx': ['Code barre', 'Code article', 'Désignation', 'PCB', 'Fournisseur']
};

// Alias pour la détection des en-têtes
const COLUMN_ALIASES = {
    'Code barre': ['code barre', 'codebarre', 'ean', 'codebar', 'code-barre'],
    'Code article': ['code article', 'codearticle', 'ref', 'reference', 'art'],
    'Désignation': ['designation', 'des', 'produit', 'libellé', 'libelle', 'article'],
    'PCB': ['pcb', 'prix', 'prix unitaire'],
    'Fournisseur': ['fournisseur', 'fourn', 'fourn.', 'supplier']
};

function detectHeaderIndices(headerRow, targetColumns) {
    const indices = {};
    const clean = headerRow.map(c => (c || '').toString().toLowerCase().trim());
    targetColumns.forEach(col => {
        const aliases = COLUMN_ALIASES[col] || [col.toLowerCase()];
        const idx = clean.findIndex(c => aliases.some(a => c.includes(a)));
        if (idx !== -1) indices[col] = idx;
    });
    return indices;
}

async function reorganizeExcel(base64, targetColumns) {
    const buffer = Buffer.from(base64, 'base64');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];

    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row) => {
        rows.push(row.values.slice(1)); // row.values[0] est undefined
    });
    if (rows.length === 0) throw new Error('Fichier vide');

    // Détecter les colonnes
    const headerRow = rows[0];
    const hasHeader = headerRow.some(cell => cell && cell.toString().toLowerCase().includes('code'));
    if (!hasHeader) throw new Error('Le fichier doit contenir une ligne d\'en-tête avec les noms de colonnes.');

    const colMap = detectHeaderIndices(headerRow, targetColumns);
    const missing = targetColumns.filter(col => !(col in colMap));
    if (missing.length > 0) throw new Error(`Colonnes manquantes : ${missing.join(', ')}`);

    // Réorganiser les lignes
    const newRows = [targetColumns]; // nouvelle en-tête
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        const newRow = targetColumns.map(col => {
            const idx = colMap[col];
            return (idx !== undefined && row[idx] !== undefined) ? row[idx] : '';
        });
        newRows.push(newRow);
    }

    // Créer un nouveau workbook
    const newWorkbook = new ExcelJS.Workbook();
    const newWs = newWorkbook.addWorksheet(ws.name);
    newRows.forEach(row => newWs.addRow(row));
    const outBuffer = await newWorkbook.xlsx.writeBuffer();
    return outBuffer.toString('base64');
}

export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Méthode non autorisée' });

    try {
        const { fileBase64, fileName, mode } = req.body;
        if (!fileBase64 || !fileName || !mode) {
            return res.status(400).json({ error: 'Paramètres manquants' });
        }

        const token = process.env.GITHUB_TOKEN;
        if (!token) {
            console.error('GITHUB_TOKEN manquant');
            return res.status(500).json({ error: 'Token GitHub non configuré' });
        }

        const repoOwner = 'hasiniainahasina1-art';
        const repoName = 'BOTAKI';
        const filePath = fileName;

        // Déterminer les colonnes cibles en fonction du fichier
        const targetColumns = FILE_STRUCTURE[fileName];
        if (!targetColumns) return res.status(400).json({ error: 'Type de fichier inconnu' });

        let finalBase64 = fileBase64;

        // Pour le mode remplacement, on réorganise toujours les colonnes
        if (mode === 'replace') {
            try {
                finalBase64 = await reorganizeExcel(fileBase64, targetColumns);
            } catch (err) {
                return res.status(400).json({ error: 'Erreur lors de la réorganisation : ' + err.message });
            }
            const content = finalBase64;
            await commitFile(token, repoOwner, repoName, filePath, content, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }
        else if (mode === 'append') {
            // Pour l'ajout, on fusionne avec l'existant (comportement inchangé)
            const existingBuffer = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingBuffer) {
                return res.status(404).json({ error: 'Fichier existant introuvable' });
            }

            // Réorganiser d'abord le nouveau fichier pour qu'il ait les mêmes colonnes dans le même ordre
            finalBase64 = await reorganizeExcel(fileBase64, targetColumns);

            const oldWorkbook = new ExcelJS.Workbook();
            await oldWorkbook.xlsx.load(existingBuffer);
            const newWorkbook = new ExcelJS.Workbook();
            await newWorkbook.xlsx.load(Buffer.from(finalBase64, 'base64'));

            const oldSheet = oldWorkbook.worksheets[0];
            const newSheet = newWorkbook.worksheets[0];

            const oldRows = [];
            oldSheet.eachRow({ includeEmpty: true }, (row) => {
                oldRows.push(row.values.slice(1));
            });
            const newRows = [];
            newSheet.eachRow({ includeEmpty: true }, (row) => {
                newRows.push(row.values.slice(1));
            });

            // Normalement les deux fichiers ont maintenant exactement les mêmes en-têtes, donc on peut concaténer les lignes de données
            const header = oldRows[0]; // on garde l'en-tête existant
            const oldBody = oldRows.slice(1);
            const newBody = newRows.slice(1); // on ignore l'en-tête du nouveau fichier

            const mergedData = [header, ...oldBody, ...newBody];

            const mergedWorkbook = new ExcelJS.Workbook();
            const ws = mergedWorkbook.addWorksheet(oldSheet.name);
            mergedData.forEach(row => ws.addRow(row));
            const outBuffer = await mergedWorkbook.xlsx.writeBuffer();
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

// --- Fonctions GitHub (inchangées) ---
async function getFileContent(token, owner, repo, path) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const resp = await fetch(url, { headers: { Authorization: `token ${token}` } });
    if (!resp.ok) return null;
    const data = await resp.json();
    return Buffer.from(data.content, 'base64');
}

async function getFileSha(token, owner, repo, path) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const resp = await fetch(url, { headers: { Authorization: `token ${token}` } });
    if (!resp.ok) return null;
    const data = await resp.json();
    return data.sha;
}

async function commitFile(token, owner, repo, path, contentBase64, message) {
    const sha = await getFileSha(token, owner, repo, path);
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const body = { message, content: contentBase64, branch: 'main' };
    if (sha) body.sha = sha;
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
