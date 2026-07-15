import ExcelJS from 'exceljs';

const FILE_STRUCTURE = {
    'database.xlsx': ['Code barre', 'Code article', 'Désignation'],
    'CADENCIER.xlsx': ['Code barre', 'Code article', 'Désignation', 'PCB', 'Fournisseur']
};

const COLUMN_ALIASES = {
    'Code barre': ['code barre', 'code-barre', 'codebarre', 'ean', 'codebar'],
    'Code article': ['code article', 'codearticle', 'ref', 'reference', 'art'],
    'Désignation': ['designation', 'désignation', 'libelle', 'libellé', 'description', 'nom', 'produit', 'design'],
    'PCB': ['pcb', 'prix unitaire', 'prix'],
    'Fournisseur': ['fournisseur', 'fourn.', 'fourn', 'supplier']
};

function normalize(str) {
    return str.toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
}

function detectHeaderIndices(headerRow, targetColumns) {
    const indices = {};
    const used = new Set();
    const clean = headerRow.map(cell => normalize(cell || ''));

    for (const col of targetColumns) {
        const aliases = COLUMN_ALIASES[col] || [col.toLowerCase()];
        let bestIdx = -1, bestScore = -1;
        for (let i = 0; i < clean.length; i++) {
            if (used.has(i)) continue;
            const cell = clean[i];
            for (const alias of aliases) {
                if (cell === alias || cell.includes(alias)) {
                    const score = alias.length;
                    if (score > bestScore) { bestScore = score; bestIdx = i; }
                    break;
                }
            }
        }
        if (bestIdx !== -1) { indices[col] = bestIdx; used.add(bestIdx); }
    }
    return indices;
}

async function reorganizeExcel(base64, targetColumns) {
    const buffer = Buffer.from(base64, 'base64');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];

    const rows = [];
    ws.eachRow({ includeEmpty: true }, row => {
        const vals = [];
        for (let i = 1; i < row.values.length; i++) vals.push(row.values[i] !== undefined ? row.values[i] : '');
        rows.push(vals);
    });
    if (rows.length === 0) throw new Error('Fichier vide');

    const header = rows[0];
    const hasHeader = header.some(cell => ['code', 'barre', 'article', 'design', 'produit', 'nom', 'libell', 'pcb', 'fourn'].some(k => normalize(cell || '').includes(k)));
    if (!hasHeader) throw new Error('Aucune ligne d\'en-tête détectée.');

    const colMap = detectHeaderIndices(header, targetColumns);
    const missing = targetColumns.filter(c => !(c in colMap));
    if (missing.length > 0) throw new Error(`Colonnes manquantes : ${missing.join(', ')}. En-têtes trouvés : ${header.join(', ')}`);

    const newRows = [targetColumns];
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        newRows.push(targetColumns.map(col => {
            const idx = colMap[col];
            return (idx !== undefined && idx < row.length) ? (row[idx] || '') : '';
        }));
    }

    const outWb = new ExcelJS.Workbook();
    const outWs = outWb.addWorksheet(ws.name);
    newRows.forEach(row => outWs.addRow(row));
    return (await outWb.xlsx.writeBuffer()).toString('base64');
}

/**
 * Fusionne les données :
 * - Identifie les produits par Code article (puis Code barre en fallback).
 * - Met à jour les colonnes vides des produits existants.
 * - Ajoute les nouveaux produits.
 */
function mergeProducts(oldRows, newRows, targetColumns) {
    const header = oldRows[0];
    const oldData = oldRows.slice(1);
    const newData = newRows.slice(1);

    const idxCodeArticle = targetColumns.indexOf('Code article');
    const idxCodeBarre = targetColumns.indexOf('Code barre');

    // Normaliser une valeur pour servir de clé
    const keyOf = (row) => {
        let key = '';
        if (idxCodeArticle !== -1) key = normalize(row[idxCodeArticle] || '');
        if (!key && idxCodeBarre !== -1) key = normalize(row[idxCodeBarre] || '');
        return key;
    };

    // Index des anciens produits
    const existingMap = new Map();
    oldData.forEach((row, i) => {
        const k = keyOf(row);
        if (k) existingMap.set(k, { row, index: i });
    });

    let added = 0, updated = 0;

    for (const newRow of newData) {
        const k = keyOf(newRow);
        if (k && existingMap.has(k)) {
            // Produit existant : remplir les vides
            const oldRow = existingMap.get(k).row;
            let changed = false;
            for (let c = 0; c < targetColumns.length; c++) {
                const oldVal = (oldRow[c] || '').toString().trim();
                const newVal = (newRow[c] || '').toString().trim();
                if (!oldVal && newVal) {
                    oldRow[c] = newRow[c];
                    changed = true;
                }
            }
            if (changed) updated++;
        } else {
            // Nouveau produit
            oldData.push(newRow);
            if (k) existingMap.set(k, { row: newRow, index: oldData.length - 1 });
            added++;
        }
    }

    return { mergedRows: [header, ...oldData], added, updated };
}

export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Méthode non autorisée' });

    try {
        const { fileBase64, fileName, mode } = req.body;
        if (!fileBase64 || !fileName || !mode) return res.status(400).json({ error: 'Paramètres manquants' });

        const token = process.env.GITHUB_TOKEN;
        if (!token) return res.status(500).json({ error: 'Token GitHub manquant' });

        const repoOwner = 'hasiniainahasina1-art', repoName = 'BOTAKI', filePath = fileName;
        const targetColumns = FILE_STRUCTURE[fileName];
        if (!targetColumns) return res.status(400).json({ error: 'Fichier inconnu' });

        if (mode === 'replace') {
            const finalB64 = await reorganizeExcel(fileBase64, targetColumns);
            await commitFile(token, repoOwner, repoName, filePath, finalB64, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }
        else if (mode === 'append') {
            const existingBuf = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingBuf) return res.status(404).json({ error: 'Fichier existant introuvable' });

            const finalB64 = await reorganizeExcel(fileBase64, targetColumns);
            const oldWb = new ExcelJS.Workbook(); await oldWb.xlsx.load(existingBuf);
            const newWb = new ExcelJS.Workbook(); await newWb.xlsx.load(Buffer.from(finalB64, 'base64'));

            const extract = (ws) => {
                const r = [];
                ws.eachRow({ includeEmpty: true }, row => {
                    const vals = [];
                    for (let i = 1; i < row.values.length; i++) vals.push(row.values[i] !== undefined ? row.values[i] : '');
                    r.push(vals);
                });
                return r;
            };

            const oldRows = extract(oldWb.worksheets[0]);
            const newRows = extract(newWb.worksheets[0]);
            const { mergedRows, added, updated } = mergeProducts(oldRows, newRows, targetColumns);

            const mergedWb = new ExcelJS.Workbook();
            const ws = mergedWb.addWorksheet(oldWb.worksheets[0].name);
            mergedRows.forEach(row => ws.addRow(row));
            const outBuf = await mergedWb.xlsx.writeBuffer();
            const content = outBuf.toString('base64');
            const msg = `Ajout ${fileName} – ${added} nouveau(x), ${updated} mis à jour`;
            await commitFile(token, repoOwner, repoName, filePath, content, msg);
            return res.status(200).json({ success: true, added, updated });
        }
        else return res.status(400).json({ error: 'Mode invalide' });
    } catch (err) {
        console.error(err);
        return res.status(500).json({ error: err.message });
    }
}

// --- Fonctions GitHub (inchangées) ---
async function getFileContent(token, owner, repo, path) {
    const resp = await fetch(`https://api.github.com/repos/${owner}/${repo}/contents/${path}`, { headers: { Authorization: `token ${token}` } });
    if (!resp.ok) return null;
    const data = await resp.json();
    return Buffer.from(data.content, 'base64');
}

async function getFileSha(token, owner, repo, path) {
    const resp = await fetch(`https://api.github.com/repos/${owner}/${repo}/contents/${path}`, { headers: { Authorization: `token ${token}` } });
    if (!resp.ok) return null;
    const data = await resp.json();
    return data.sha;
}

async function commitFile(token, owner, repo, path, contentB64, message) {
    const sha = await getFileSha(token, owner, repo, path);
    const body = { message, content: contentB64, branch: 'main' };
    if (sha) body.sha = sha;
    const resp = await fetch(`https://api.github.com/repos/${owner}/${repo}/contents/${path}`, {
        method: 'PUT', headers: { Authorization: `token ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(body)
    });
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.message); }
}
