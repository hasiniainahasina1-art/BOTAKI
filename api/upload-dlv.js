import ExcelJS from 'exceljs';

const REPO_OWNER = 'hasiniainahasina1-art';
const REPO_NAME  = 'BOTAKI';
const SHEETS     = ['Épicerie', 'Boisson', 'DPH'];
const COLUMNS    = ['Code Barre', 'Code Article', 'Désignation', 'Prix Normal', 'Prix Promo', 'Quantité', 'Dlv'];

// ── Helpers GitHub ──────────────────────────────────────────────
async function getFile(token, path) {
    const resp = await fetch(`https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${path}`, {
        headers: { Authorization: `token ${token}`, Accept: 'application/vnd.github.v3+json' }
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    return { sha: data.sha, buffer: Buffer.from(data.content, 'base64') };
}

async function commitFile(token, path, contentB64, sha, message) {
    const body = { message, content: contentB64, branch: 'main' };
    if (sha) body.sha = sha;
    const resp = await fetch(`https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${path}`, {
        method: 'PUT',
        headers: { Authorization: `token ${token}`, 'Content-Type': 'application/json', Accept: 'application/vnd.github.v3+json' },
        body: JSON.stringify(body)
    });
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.message); }
}

// ── Créer un classeur vide avec les 3 feuilles ───────────────────
async function creerClasseurVide() {
    const wb = new ExcelJS.Workbook();
    for (const sheetName of SHEETS) {
        const ws = wb.addWorksheet(sheetName);
        const headerRow = ws.addRow(COLUMNS);
        headerRow.eachCell(cell => {
            cell.font  = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1e2f6b' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });
        ws.columns = COLUMNS.map((col, i) => ({ key: col, width: [18, 18, 32, 16, 16, 12, 16][i] }));
    }
    return wb;
}

// ── Handler principal ────────────────────────────────────────────
export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Méthode non autorisée' });

    const { emplacement, categorie, ligne, mode } = req.body;

    // Validation
    if (!emplacement || !['KIBO4', 'KIBO2'].includes(emplacement))
        return res.status(400).json({ error: 'Emplacement invalide (KIBO4 ou KIBO2)' });
    if (!categorie || !SHEETS.includes(categorie))
        return res.status(400).json({ error: `Catégorie invalide (${SHEETS.join(', ')})` });

    const token = process.env.GITHUB_TOKEN;
    if (!token) return res.status(500).json({ error: 'GITHUB_TOKEN manquant' });

    const filePath = `dlv_${emplacement}.xlsx`;

    try {
        // ── Mode READ : lire le fichier pour la visualisation ──────
        if (mode === 'read') {
            const existing = await getFile(token, filePath);
            if (!existing) return res.status(200).json({ exists: false, sheets: {} });

            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(existing.buffer);

            const result = {};
            for (const sheetName of SHEETS) {
                const ws = wb.getWorksheet(sheetName);
                if (!ws) { result[sheetName] = []; continue; }
                const rows = [];
                ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                    const vals = [];
                    for (let i = 1; i <= COLUMNS.length; i++) {
                        const cell = row.getCell(i);
                        let v = cell.value;
                        if (v && typeof v === 'object' && v.result !== undefined) v = v.result;
                        if (v instanceof Date) {
                            const d = v;
                            v = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
                        }
                        vals.push(v !== null && v !== undefined ? String(v) : '');
                    }
                    rows.push(vals);
                });
                result[sheetName] = rows;
            }
            return res.status(200).json({ exists: true, sheets: result });
        }

        // ── Mode WRITE : insérer une nouvelle ligne ────────────────
        if (mode === 'write') {
            if (!ligne) return res.status(400).json({ error: 'Paramètre ligne manquant' });

            let wb = new ExcelJS.Workbook();
            let sha = null;

            const existing = await getFile(token, filePath);
            if (existing) {
                await wb.xlsx.load(existing.buffer);
                sha = existing.sha;
            } else {
                wb = await creerClasseurVide();
            }

            const ws = wb.getWorksheet(categorie) || wb.addWorksheet(categorie);

            // S'assurer que l'en-tête existe
            const firstRow = ws.getRow(1);
            const firstCell = firstRow.getCell(1).value;
            if (!firstCell || firstCell.toString().toLowerCase() !== 'code barre') {
                firstRow.values = COLUMNS;
                firstRow.eachCell(cell => {
                    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' } };
                    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1e2f6b' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                });
            }

            // Insérer en ligne 2 (après l'en-tête) → date la plus récente en haut
            // Décaler toutes les lignes existantes d'une position vers le bas
            const totalRows = ws.rowCount;
            for (let r = totalRows; r >= 2; r--) {
                const srcRow = ws.getRow(r);
                const dstRow = ws.getRow(r + 1);
                dstRow.values = srcRow.values;
                dstRow.eachCell(cell => { cell.alignment = { horizontal: 'center', vertical: 'middle' }; });
            }

            // Écrire la nouvelle ligne en position 2
            const newRow = ws.getRow(2);
            newRow.values = [
                ligne.codeBarre    || '',
                ligne.codeArticle  || '',
                ligne.designation  || '',
                ligne.prixNormal   || '',
                '',                         // Prix Promo vide
                ligne.quantite     || '',
                ligne.dlv          || ''
            ];
            newRow.eachCell(cell => { cell.alignment = { horizontal: 'center', vertical: 'middle' }; });

            // Largeurs colonnes
            ws.columns = COLUMNS.map((col, i) => ({ width: [18, 18, 32, 16, 16, 12, 16][i] }));

            const buffer = await wb.xlsx.writeBuffer();
            const contentB64 = buffer.toString('base64');
            const msg = `Ajout dlv ${emplacement}/${categorie} – ${ligne.designation || ligne.codeBarre}`;
            await commitFile(token, filePath, contentB64, sha, msg);

            return res.status(200).json({ success: true });
        }

        // ── Mode SAVE : sauvegarder modifications depuis visualisation ──
        if (mode === 'save') {
            const { sheets } = req.body;
            if (!sheets) return res.status(400).json({ error: 'Paramètre sheets manquant' });

            const existing = await getFile(token, filePath);
            let sha = existing ? existing.sha : null;

            let wb = new ExcelJS.Workbook();
            if (existing) await wb.xlsx.load(existing.buffer);
            else wb = await creerClasseurVide();

            for (const sheetName of SHEETS) {
                let ws = wb.getWorksheet(sheetName);
                if (!ws) ws = wb.addWorksheet(sheetName);

                const rows = sheets[sheetName] || [];
                // Vider la feuille
                ws.spliceRows(1, ws.rowCount);

                // Réécrire toutes les lignes
                rows.forEach((rowData, idx) => {
                    const row = ws.addRow(rowData);
                    if (idx === 0) {
                        row.eachCell(cell => {
                            cell.font  = { bold: true, color: { argb: 'FFFFFFFF' } };
                            cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1e2f6b' } };
                            cell.alignment = { horizontal: 'center', vertical: 'middle' };
                        });
                    } else {
                        row.eachCell(cell => { cell.alignment = { horizontal: 'center', vertical: 'middle' }; });
                    }
                });
                ws.columns = COLUMNS.map((col, i) => ({ width: [18, 18, 32, 16, 16, 12, 16][i] }));
            }

            const buffer = await wb.xlsx.writeBuffer();
            const contentB64 = buffer.toString('base64');
            await commitFile(token, filePath, contentB64, sha, `Mise à jour manuelle dlv_${emplacement}.xlsx`);
            return res.status(200).json({ success: true });
        }

        return res.status(400).json({ error: 'Mode invalide (read, write, save)' });

    } catch (err) {
        console.error('upload-dlv error:', err);
        return res.status(500).json({ error: err.message });
    }
}
