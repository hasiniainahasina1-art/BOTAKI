import ExcelJS from 'exceljs';

const REPO_OWNER = 'Finoana123';
const REPO_NAME = 'Dlv';
const VALID_FOLDERS = ['KIBO4', 'KIBO2'];
const VALID_SHEETS = ['Épicerie', 'Boisson', 'DPH'];

// Colonnes de données (sans Date de mise à jour)
const HEADERS = [
    'Code Article',
    'Désignation',
    'Dlv',
    'Quantité',
    'Prix Normal',
    'Prix promo'
];
const COL_COUNT = 6;

function getFilePath(folder) {
    return `${folder}/dlv-prix.xlsx`;
}

async function getFileFromGitHub(token, path) {
    const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${path}`;
    const resp = await fetch(url, {
        headers: {
            Authorization: `token ${token}`,
            Accept: 'application/vnd.github.v3+json'
        }
    });
    if (resp.status === 404) return null;
    if (!resp.ok) {
        const err = await resp.json();
        throw new Error(err.message || 'Erreur lecture GitHub');
    }
    const data = await resp.json();
    return {
        content: Buffer.from(data.content, 'base64'),
        sha: data.sha
    };
}

async function putFileToGitHub(token, path, contentBase64, sha, message) {
    const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${path}`;

    // Toujours récupérer le SHA actuel si le fichier existe (évite l'erreur "sha wasn't supplied")
    let currentSha = sha || null;
    try {
        const existing = await getFileFromGitHub(token, path);
        if (existing && existing.sha) currentSha = existing.sha;
    } catch (_) { /* ignore */ }

    const body = {
        message,
        content: contentBase64,
        branch: 'main'
    };
    if (currentSha) body.sha = currentSha;

    const resp = await fetch(url, {
        method: 'PUT',
        headers: {
            Authorization: `token ${token}`,
            'Content-Type': 'application/json',
            Accept: 'application/vnd.github.v3+json'
        },
        body: JSON.stringify(body)
    });
    if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        const detail = err.message || resp.statusText || 'Erreur écriture GitHub';
        // Messages plus clairs
        if (resp.status === 404) {
            throw new Error('Dépôt ou fichier introuvable (404). Vérifiez REPO Finoana123/Dlv et le token GITHUB_TOKEN.');
        }
        if (resp.status === 401 || resp.status === 403) {
            throw new Error('Token GitHub refusé (' + resp.status + '). Vérifiez GITHUB_TOKEN sur Vercel (droit repo).');
        }
        throw new Error(detail);
    }
    return await resp.json();
}

function argbToHex(argb) {
    if (!argb) return '';
    const s = String(argb).replace(/^FF/i, '').replace(/^#/, '');
    if (s.length === 6) return '#' + s.toLowerCase();
    if (s.length === 8) return '#' + s.slice(2).toLowerCase();
    return '';
}

function hexToArgb(hex) {
    if (!hex) return null;
    const s = String(hex).replace(/^#/, '').toUpperCase();
    if (s.length === 6) return 'FF' + s;
    if (s.length === 8) return s;
    return null;
}

/** Convertit toute date Excel/ISO en texte JJ/MM/AAAA */
function normalizeDlv(val) {
    if (val === null || val === undefined || val === '') return '';
    if (val instanceof Date) {
        const d = val.getDate().toString().padStart(2, '0');
        const m = (val.getMonth() + 1).toString().padStart(2, '0');
        const y = val.getFullYear();
        return `${d}/${m}/${y}`;
    }
    // ExcelJS sometimes gives serial number
    if (typeof val === 'number' && val > 20000 && val < 100000) {
        const utc = new Date(Date.UTC(1899, 11, 30) + val * 86400000);
        const d = utc.getUTCDate().toString().padStart(2, '0');
        const m = (utc.getUTCMonth() + 1).toString().padStart(2, '0');
        const y = utc.getUTCFullYear();
        return `${d}/${m}/${y}`;
    }
    const s = String(val).trim();
    // ISO YYYY-MM-DD
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[3]}/${m[2]}/${m[1]}`;
    // already JJ/MM/AAAA
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return `${m[1].padStart(2,'0')}/${m[2].padStart(2,'0')}/${m[3]}`;
    return s;
}


async function parseWorkbook(buffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const result = {
        updateDate: '',
        sheets: {}
    };

    for (const sheetName of VALID_SHEETS) {
        const ws = wb.getWorksheet(sheetName);
        const rows = [];

        if (ws) {
            // Row 1 = Date de mise à jour globale
            const dateCell = ws.getCell('A1').value;
            if (dateCell && !result.updateDate) {
                result.updateDate = String(dateCell).replace(/^Date de mise à jour\s*:\s*/i, '').trim();
            }

            // Detect header layout (ancien format avec Code Barre vs nouveau)
            const headerRow = ws.getRow(2);
            const headers = [];
            for (let c = 1; c <= 10; c++) {
                const v = headerRow.getCell(c).value;
                headers.push(v ? String(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '') : '');
            }
            const findCol = (...names) => {
                for (const n of names) {
                    const i = headers.findIndex(h => h.includes(n));
                    if (i >= 0) return i + 1;
                }
                return -1;
            };
            // Nouveau format par défaut : Code Article, Désignation, Dlv, Quantité, Prix Normal, Prix promo
            let colArticle = findCol('code article', 'codearticle', 'article');
            let colDesign = findCol('designation', 'libelle', 'nom');
            let colDlv = findCol('dlv', 'date limite');
            let colQte = findCol('quantite', 'qte', 'qty');
            let colPrixN = findCol('prix normal', 'prix normale', 'prix');
            let colPrixP = findCol('prix promo', 'promo');
            // Ancien format : Code Barre, Code article, Désignation, Prix Normale, Prix promo, Quantité, Dlv
            if (colArticle < 0 && headers[0] && headers[0].includes('code barre')) {
                colArticle = 2; colDesign = 3; colPrixN = 4; colPrixP = 5; colQte = 6; colDlv = 7;
            }
            if (colArticle < 0) colArticle = 1;
            if (colDesign < 0) colDesign = 2;
            if (colDlv < 0) colDlv = 3;
            if (colQte < 0) colQte = 4;
            if (colPrixN < 0) colPrixN = 5;
            if (colPrixP < 0) colPrixP = 6;

            const maxCol = Math.max(colArticle, colDesign, colDlv, colQte, colPrixN, colPrixP, 6);

            // Data starts at row 3 (row 2 = headers)
            ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                if (rowNumber < 3) return;
                const read = (col) => {
                    if (col < 1) return '';
                    const cell = row.getCell(col);
                    let val = cell.value;
                    if (val && typeof val === 'object' && val.result !== undefined) val = val.result;
                    if (val && typeof val === 'object' && val.text) val = val.text;
                    return val !== null && val !== undefined ? String(val) : '';
                };
                const codeArticle = read(colArticle);
                const designation = read(colDesign);
                // Lire la valeur brute pour DLV puis normaliser
                let dlvRaw = null;
                if (colDlv > 0) {
                    const cell = row.getCell(colDlv);
                    dlvRaw = cell.value;
                    if (dlvRaw && typeof dlvRaw === 'object' && dlvRaw.result !== undefined) dlvRaw = dlvRaw.result;
                }
                const dlv = normalizeDlv(dlvRaw);
                const quantite = read(colQte);
                const prixNormale = read(colPrixN);
                const prixPromo = read(colPrixP);
                if (![codeArticle, designation, dlv, quantite, prixNormale, prixPromo].some(v => v && v.trim())) return;

                const colors = [];
                // Couleurs dans l'ordre des nouvelles colonnes
                const colorCols = [colArticle, colDesign, colDlv, colQte, colPrixN, colPrixP];
                for (const c of colorCols) {
                    let hex = '';
                    try {
                        const cell = row.getCell(c);
                        const fill = cell.fill;
                        if (fill && fill.type === 'pattern' && fill.fgColor && fill.fgColor.argb) {
                            const h = argbToHex(fill.fgColor.argb);
                            if (h && h.toLowerCase() !== '#1e2f6b') hex = h;
                        }
                    } catch (_) { /* ignore */ }
                    colors.push(hex);
                }
                rows.push({
                    codeArticle: codeArticle || '',
                    designation: designation || '',
                    dlv: dlv || '',
                    quantite: quantite || '',
                    prixNormale: prixNormale || '',
                    prixPromo: prixPromo || '',
                    _cellColors: colors
                });
            });
        }
        result.sheets[sheetName] = rows;
    }
    return result;
}

async function buildWorkbook(data) {
    const wb = new ExcelJS.Workbook();

    const headerFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E2F6B' }
    };
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    const thinBorder = {
        top: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        left: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        bottom: { style: 'thin', color: { argb: 'FFCCCCCC' } },
        right: { style: 'thin', color: { argb: 'FFCCCCCC' } }
    };

    const now = new Date();
    const dateStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    function parseDlvTs(str) {
        if (!str) return 0;
        const m = String(str).match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (!m) return 0;
        return new Date(+m[3], +m[2] - 1, +m[1]).getTime() || 0;
    }

    for (const sheetName of VALID_SHEETS) {
        const ws = wb.addWorksheet(sheetName);
        const rows = (data.sheets[sheetName] || []).slice();
        // Date DLV la plus proche (ancienne) en haut
        rows.sort((a, b) => (parseDlvTs(a.dlv) || Number.MAX_SAFE_INTEGER) - (parseDlvTs(b.dlv) || Number.MAX_SAFE_INTEGER));

        // Row 1: Date de mise à jour globale uniquement
        ws.mergeCells('A1:F1');
        ws.getCell('A1').value = `Date de mise à jour : ${dateStr}`;
        ws.getCell('A1').font = { bold: true, size: 12, color: { argb: 'FF1E2F6B' } };
        ws.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle' };
        ws.getRow(1).height = 22;

        // Row 2: Headers
        HEADERS.forEach((h, i) => {
            const cell = ws.getCell(2, i + 1);
            cell.value = h;
            cell.fill = headerFill;
            cell.font = headerFont;
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = thinBorder;
        });
        ws.getRow(2).height = 20;

        // Data rows (avec couleurs éventuelles)
        rows.forEach((r, idx) => {
            const rowNum = idx + 3;
            const values = [
                r.codeArticle || '',
                r.designation || '',
                normalizeDlv(r.dlv) || '',
                r.quantite || '',
                r.prixNormale || '',
                r.prixPromo || ''
            ];
            const colors = Array.isArray(r._cellColors) ? r._cellColors : [];
            values.forEach((v, c) => {
                const cell = ws.getCell(rowNum, c + 1);
                cell.value = v;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = thinBorder;
                // Colonne Dlv (index 2) toujours en texte JJ/MM/AAAA
                if (c === 2) cell.numFmt = '@';
                const argb = hexToArgb(colors[c]);
                if (argb) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb }
                    };
                }
            });
            ws.getRow(rowNum).height = 18;
        });

        const widths = [14, 45, 12, 10, 14, 12];
        widths.forEach((w, i) => {
            ws.getColumn(i + 1).width = w;
        });

        ws.views = [{ state: 'frozen', ySplit: 2 }];
    }

    const buffer = await wb.xlsx.writeBuffer();
    return Buffer.from(buffer);
}

export default async function handler(req, res) {
    const token = process.env.GITHUB_TOKEN;
    if (!token) {
        return res.status(500).json({ error: 'Token GitHub non configuré (GITHUB_TOKEN)' });
    }

    try {
        if (req.method === 'GET') {
            const folder = (req.query.folder || '').toUpperCase();
            if (!VALID_FOLDERS.includes(folder)) {
                return res.status(400).json({ error: 'Dossier invalide. Utilisez KIBO4 ou KIBO2.' });
            }

            const path = getFilePath(folder);
            const file = await getFileFromGitHub(token, path);

            if (!file) {
                return res.status(200).json({
                    success: true,
                    folder,
                    updateDate: '',
                    sheets: {
                        'Épicerie': [],
                        'Boisson': [],
                        'DPH': []
                    },
                    sha: null
                });
            }

            const data = await parseWorkbook(file.content);
            return res.status(200).json({
                success: true,
                folder,
                updateDate: data.updateDate,
                sheets: data.sheets,
                sha: file.sha
            });
        }

        if (req.method === 'POST') {
            const { folder, sheets, sha } = req.body || {};

            if (!folder || !VALID_FOLDERS.includes(folder.toUpperCase())) {
                return res.status(400).json({ error: 'Dossier invalide. Utilisez KIBO4 ou KIBO2.' });
            }
            if (!sheets || typeof sheets !== 'object') {
                return res.status(400).json({ error: 'Paramètre sheets manquant ou invalide.' });
            }

            const path = getFilePath(folder.toUpperCase());
            const buffer = await buildWorkbook({ sheets });
            const contentBase64 = buffer.toString('base64');

            const result = await putFileToGitHub(
                token,
                path,
                contentBase64,
                sha || null,
                `Mise à jour DLV/Prix ${folder.toUpperCase()}`
            );

            return res.status(200).json({
                success: true,
                folder: folder.toUpperCase(),
                sha: result.content?.sha || null,
                url: result.content?.html_url || null
            });
        }

        return res.status(405).json({ error: 'Méthode non autorisée. Utilisez GET ou POST.' });
    } catch (error) {
        console.error('[dlv-prix]', error);
        return res.status(500).json({ error: error.message || 'Erreur serveur' });
    }
}
