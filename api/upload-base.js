import ExcelJS from 'exceljs';

// Colonnes attendues pour chaque type de fichier
const FILE_STRUCTURE = {
    'database.xlsx': ['Code barre', 'Code article', 'Désignation'],
    'CADENCIER.xlsx': ['Code barre', 'Code article', 'Désignation', 'PCB', 'Fournisseur']
};

// Alias très précis pour éviter les confusions
const COLUMN_ALIASES = {
    'Code barre': ['code barre', 'code-barre', 'codebarre', 'ean', 'codebar'],
    'Code article': ['code article', 'codearticle', 'ref', 'reference', 'art'],   // 'art' est volontairement gardé court pour ne pas être confondu
    'Désignation': ['designation', 'désignation', 'libelle', 'libellé', 'description', 'nom', 'produit', 'design'],
    'PCB': ['pcb', 'prix unitaire', 'prix'],
    'Fournisseur': ['fournisseur', 'fourn.', 'fourn', 'supplier']
};

/**
 * Nettoie une chaîne : minuscule, sans accent, sans espaces superflus.
 */
function normalize(str) {
    return str
        .toString()
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .trim();
}

/**
 * Détecte les indices des colonnes cibles.
 * - Une colonne source ne peut être attribuée qu'à UNE SEULE colonne cible.
 * - Les correspondances les plus longues sont privilégiées.
 */
function detectHeaderIndices(headerRow, targetColumns) {
    const indices = {};
    const usedSourceIndices = new Set();
    const clean = headerRow.map(cell => normalize(cell || ''));

    for (const col of targetColumns) {
        const aliases = COLUMN_ALIASES[col] || [col.toLowerCase()];
        let bestIdx = -1;
        let bestScore = -1;

        for (let i = 0; i < clean.length; i++) {
            if (usedSourceIndices.has(i)) continue; // déjà prise

            const cell = clean[i];
            for (const alias of aliases) {
                // On vérifie si la cellule contient l'alias (correspondance exacte ou en tant que mot)
                if (cell === alias || cell.includes(alias)) {
                    const score = alias.length;
                    if (score > bestScore) {
                        bestScore = score;
                        bestIdx = i;
                    }
                    break; // on passe à la colonne suivante
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
 * Réorganise un fichier Excel (base64) pour qu'il ait exactement les colonnes cibles.
 */
async function reorganizeExcel(base64, targetColumns) {
    const buffer = Buffer.from(base64, 'base64');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const ws = workbook.worksheets[0];

    const rows = [];
    ws.eachRow({ includeEmpty: true }, (row) => {
        const vals = [];
        for (let i = 1; i < row.values.length; i++) {
            vals.push(row.values[i] !== undefined ? row.values[i] : '');
        }
        rows.push(vals);
    });
    if (rows.length === 0) throw new Error('Fichier vide');

    const headerRow = rows[0];
    const hasHeader = headerRow.some(cell => {
        const v = normalize(cell || '');
        return ['code', 'barre', 'article', 'design', 'produit', 'nom', 'libell', 'pcb', 'fourn'].some(k => v.includes(k));
    });
    if (!hasHeader) throw new Error('Le fichier doit contenir une ligne d\'en-tête.');

    const colMap = detectHeaderIndices(headerRow, targetColumns);
    const missing = targetColumns.filter(col => !(col in colMap));
    if (missing.length > 0) {
        const found = headerRow.join(', ');
        throw new Error(`Colonnes manquantes : ${missing.join(', ')}. En-têtes trouvés : ${found}`);
    }

    // Réorganiser
    const newRows = [targetColumns];
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        const newRow = targetColumns.map(col => {
            const idx = colMap[col];
            return (idx !== undefined && idx < row.length) ? (row[idx] || '') : '';
        });
        newRows.push(newRow);
    }

    const newWorkbook = new ExcelJS.Workbook();
    const newWs = newWorkbook.addWorksheet(ws.name);
    newRows.forEach(row => newWs.addRow(row));
    const outBuffer = await newWorkbook.xlsx.writeBuffer();
    return outBuffer.toString('base64');
}

// --- Handler principal ---
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

        // --- REMPLACEMENT ---
        if (mode === 'replace') {
            const finalBase64 = await reorganizeExcel(fileBase64, targetColumns);
            await commitFile(token, repoOwner, repoName, filePath, finalBase64, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }
        // --- AJOUT ---
        else if (mode === 'append') {
            const existingBuffer = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingBuffer) {
                return res.status(404).json({ error: 'Fichier existant introuvable' });
            }

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

// --- Fonctions GitHub ---
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
