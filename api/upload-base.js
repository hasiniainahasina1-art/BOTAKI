import ExcelJS from 'exceljs';

// Colonnes attendues pour chaque type de fichier
const FILE_STRUCTURE = {
    'database.xlsx': ['Code barre', 'Code article', 'Désignation'],
    'CADENCIER.xlsx': ['Code barre', 'Code article', 'Désignation', 'PCB', 'Fournisseur']
};

// Alias pour la détection des en-têtes
const COLUMN_ALIASES = {
    'Code barre': ['code barre', 'code-barre', 'codebarre', 'ean', 'codebar'],
    'Code article': ['code article', 'codearticle', 'ref', 'reference', 'art'],
    'Désignation': ['designation', 'désignation', 'libelle', 'libellé', 'description', 'nom', 'produit', 'design'],
    'PCB': ['pcb', 'prix unitaire', 'prix'],
    'Fournisseur': ['fournisseur', 'fourn.', 'fourn', 'supplier']
};

/**
 * Nettoie une chaîne : minuscule, sans accent.
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
 * Détecte les indices des colonnes cibles dans l'en-tête.
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
            if (usedSourceIndices.has(i)) continue;
            const cell = clean[i];
            for (const alias of aliases) {
                if (cell === alias || cell.includes(alias)) {
                    const score = alias.length;
                    if (score > bestScore) {
                        bestScore = score;
                        bestIdx = i;
                    }
                    break;
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

/**
 * Fusionne les lignes du nouveau fichier dans les anciennes.
 * - indexCol : le nom de la colonne servant d'identifiant ('Code article' par défaut, sinon 'Code barre')
 * - Retourne { mergedRows, addedCount }
 */
function mergeProducts(oldRows, newRows, targetColumns) {
    // oldRows[0] et newRows[0] sont les en-têtes (identiques)
    const header = oldRows[0];
    const oldData = oldRows.slice(1);
    const newData = newRows.slice(1);

    // Trouver l'index de la colonne d'identification
    let idColIndex = targetColumns.indexOf('Code article');
    if (idColIndex === -1) idColIndex = targetColumns.indexOf('Code barre');
    // Colonne de secours : Code barre
    let idColIndex2 = targetColumns.indexOf('Code barre');
    if (idColIndex2 === idColIndex) idColIndex2 = -1;

    // Créer un dictionnaire des lignes existantes indexées par l'identifiant
    const existingMap = new Map(); // clé -> objet { row, index }
    oldData.forEach((row, idx) => {
        let key = (row[idColIndex] || '').toString().trim();
        if (!key && idColIndex2 !== -1) {
            key = (row[idColIndex2] || '').toString().trim();
        }
        if (key) {
            // Si plusieurs lignes ont la même clé, on ne garde que la première (on pourrait fusionner aussi)
            if (!existingMap.has(key)) {
                existingMap.set(key, { row, index: idx });
            }
        }
    });

    let addedCount = 0;

    // Parcourir les nouvelles lignes
    for (const newRow of newData) {
        let key = (newRow[idColIndex] || '').toString().trim();
        if (!key && idColIndex2 !== -1) {
            key = (newRow[idColIndex2] || '').toString().trim();
        }

        if (key && existingMap.has(key)) {
            // Produit existant : compléter les colonnes vides
            const existing = existingMap.get(key);
            const oldRow = existing.row;
            for (let c = 0; c < targetColumns.length; c++) {
                const oldVal = (oldRow[c] || '').toString().trim();
                const newVal = (newRow[c] || '').toString().trim();
                if (!oldVal && newVal) {
                    oldRow[c] = newRow[c]; // compléter avec la nouvelle valeur
                }
            }
        } else {
            // Nouveau produit
            oldData.push(newRow);
            if (key) {
                existingMap.set(key, { row: newRow, index: oldData.length - 1 });
            }
            addedCount++;
        }
    }

    return {
        mergedRows: [header, ...oldData],
        addedCount
    };
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

        // --- Mode REMPLACEMENT ---
        if (mode === 'replace') {
            const finalBase64 = await reorganizeExcel(fileBase64, targetColumns);
            await commitFile(token, repoOwner, repoName, filePath, finalBase64, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }
        // --- Mode AJOUT (fusion intelligente) ---
        else if (mode === 'append') {
            const existingBuffer = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingBuffer) {
                return res.status(404).json({ error: 'Fichier existant introuvable' });
            }

            // Réorganiser le nouveau fichier
            const finalBase64 = await reorganizeExcel(fileBase64, targetColumns);

            // Charger les deux fichiers
            const oldWorkbook = new ExcelJS.Workbook();
            await oldWorkbook.xlsx.load(existingBuffer);
            const newWorkbook = new ExcelJS.Workbook();
            await newWorkbook.xlsx.load(Buffer.from(finalBase64, 'base64'));

            // Extraire les lignes sous forme de tableaux
            const extractRows = (worksheet) => {
                const rows = [];
                worksheet.eachRow({ includeEmpty: true }, (row) => {
                    const vals = [];
                    for (let i = 1; i < row.values.length; i++) {
                        vals.push(row.values[i] !== undefined ? row.values[i] : '');
                    }
                    rows.push(vals);
                });
                return rows;
            };

            const oldRows = extractRows(oldWorkbook.worksheets[0]);
            const newRows = extractRows(newWorkbook.worksheets[0]);

            // Fusionner
            const { mergedRows, addedCount } = mergeProducts(oldRows, newRows, targetColumns);

            // Écrire le fichier fusionné
            const mergedWorkbook = new ExcelJS.Workbook();
            const ws = mergedWorkbook.addWorksheet(oldWorkbook.worksheets[0].name);
            mergedRows.forEach(row => ws.addRow(row));
            const outBuffer = await mergedWorkbook.xlsx.writeBuffer();
            const content = outBuffer.toString('base64');
            await commitFile(token, repoOwner, repoName, filePath, content, `Ajout de produits (${addedCount} nouveau(x)) dans ${fileName}`);

            return res.status(200).json({ success: true, added: addedCount });
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
