import ExcelJS from 'exceljs';

// Définition des colonnes attendues pour chaque type de fichier
const FILE_STRUCTURE = {
    'database.xlsx': ['Code barre', 'Code article', 'Désignation'],
    'CADENCIER.xlsx': ['Code barre', 'Code article', 'Désignation', 'PCB', 'Fournisseur']
};

// Alias pour la détection des en-têtes (ordonnés du plus précis au moins précis)
const COLUMN_ALIASES = {
    'Code barre': ['code barre', 'code-barre', 'codebarre', 'ean', 'codebar'],
    'Code article': ['code article', 'codearticle', 'ref', 'reference'],
    'Désignation': ['designation', 'libellé', 'libelle', 'des', 'produit', 'article'],
    'PCB': ['pcb', 'prix unitaire', 'prix'],
    'Fournisseur': ['fournisseur', 'fourn.', 'fourn', 'supplier']
};

/**
 * Détecte les indices des colonnes cibles dans une ligne d'en-tête.
 * Garantit qu'une colonne source ne peut être attribuée qu'à une seule colonne cible.
 */
function detectHeaderIndices(headerRow, targetColumns) {
    const indices = {};
    const usedSourceIndices = new Set();
    const clean = headerRow.map(c => (c || '').toString().toLowerCase().trim());

    // Pour chaque colonne cible, on cherche la meilleure correspondance non déjà utilisée
    for (const col of targetColumns) {
        const aliases = COLUMN_ALIASES[col] || [col.toLowerCase()];
        let bestIdx = -1;
        let bestScore = -1;

        for (let i = 0; i < clean.length; i++) {
            if (usedSourceIndices.has(i)) continue; // déjà prise
            const cell = clean[i];
            for (const alias of aliases) {
                if (cell.includes(alias)) {
                    // score = longueur de l'alias (préfère les correspondances plus longues)
                    const score = alias.length;
                    if (score > bestScore) {
                        bestScore = score;
                        bestIdx = i;
                    }
                    break; // on ne teste pas les autres alias pour cette cellule si on a un match
                }
            }
        }

        if (bestIdx !== -1) {
            indices[col] = bestIdx;
            usedSourceIndices.add(bestIdx);
        }
    }

    return indices;
}

/**
 * Réorganise un fichier Excel (en base64) pour qu'il ait exactement les colonnes
 * cibles dans l'ordre standard.
 */
async function reorganizeExcel(base64, targetColumns) {
    const buffer = Buffer.from(base64, 'base64');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];

    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row) => {
        // row.values est un tableau où l'index 1 correspond à la colonne A, etc.
        // On récupère toutes les valeurs à partir de l'index 1
        const vals = [];
        for (let i = 1; i < row.values.length; i++) {
            vals.push(row.values[i] !== undefined ? row.values[i] : '');
        }
        rows.push(vals);
    });
    if (rows.length === 0) throw new Error('Fichier vide');

    const headerRow = rows[0];
    const hasHeader = headerRow.some(cell => {
        const val = (cell || '').toString().toLowerCase().trim();
        return ['code', 'barre', 'article', 'designation', 'pcb', 'fournisseur'].some(key => val.includes(key));
    });
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
            return (idx !== undefined && idx < row.length) ? (row[idx] || '') : '';
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

        const targetColumns = FILE_STRUCTURE[fileName];
        if (!targetColumns) return res.status(400).json({ error: 'Type de fichier inconnu' });

        // --- Mode REMPLACEMENT : on réorganise toujours ---
        if (mode === 'replace') {
            const finalBase64 = await reorganizeExcel(fileBase64, targetColumns);
            await commitFile(token, repoOwner, repoName, filePath, finalBase64, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }

        // --- Mode AJOUT : fusion avec l'existant ---
        else if (mode === 'append') {
            const existingBuffer = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingBuffer) {
                return res.status(404).json({ error: 'Fichier existant introuvable' });
            }

            // Réorganiser d'abord le nouveau fichier pour qu'il ait les mêmes colonnes dans le même ordre
            const finalBase64 = await reorganizeExcel(fileBase64, targetColumns);

            const oldWorkbook = new ExcelJS.Workbook();
            await oldWorkbook.xlsx.load(existingBuffer);
            const newWorkbook = new ExcelJS.Workbook();
            await newWorkbook.xlsx.load(Buffer.from(finalBase64, 'base64'));

            const oldSheet = oldWorkbook.worksheets[0];
            const newSheet = newWorkbook.worksheets[0];

            const oldRows = [];
            oldSheet.eachRow({ includeEmpty: true }, (row) => {
                const vals = [];
                for (let i = 1; i < row.values.length; i++) vals.push(row.values[i] !== undefined ? row.values[i] : '');
                oldRows.push(vals);
            });
            const newRows = [];
            newSheet.eachRow({ includeEmpty: true }, (row) => {
                const vals = [];
                for (let i = 1; i < row.values.length; i++) vals.push(row.values[i] !== undefined ? row.values[i] : '');
                newRows.push(vals);
            });

            // Les deux fichiers ont maintenant la même en-tête (targetColumns)
            const header = oldRows[0];
            const oldBody = oldRows.slice(1);
            const newBody = newRows.slice(1);

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
