import ExcelJS from 'exceljs';

const REPO_OWNER = 'hasiniainahasina1-art';
const REPO_NAME = 'BOTAKI';
const VALID_FOLDERS = ['KIBO4', 'KIBO2'];
const VALID_SHEETS = ['Épicerie', 'Boisson', 'DPH'];

// Colonnes de données (sans Date de mise à jour)
const HEADERS = [
    'Code Barre',
    'Code article',
    'Désignation',
    'Prix Normale',
    'Prix promo',
    'Quantité',
    'Dlv'
];

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
    const body = {
        message,
        content: contentBase64,
        branch: 'main'
    };
    if (sha) body.sha = sha;

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
        const err = await resp.json();
        throw new Error(err.message || 'Erreur écriture GitHub');
    }
    return await resp.json();
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

            // Data starts at row 3 (row 2 = headers)
            ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                if (rowNumber < 3) return;
                const values = [];
                for (let c = 1; c <= 7; c++) {
                    const cell = row.getCell(c);
                    let val = cell.value;
                    if (val && typeof val === 'object' && val.result !== undefined) val = val.result;
                    if (val && typeof val === 'object' && val.text) val = val.text;
                    values.push(val !== null && val !== undefined ? String(val) : '');
                }
                if (values.every(v => !v || !String(v).trim())) return;
                rows.push({
                    codeBarre: values[0] || '',
                    codeArticle: values[1] || '',
                    designation: values[2] || '',
                    prixNormale: values[3] || '',
                    prixPromo: values[4] || '',
                    quantite: values[5] || '',
                    dlv: values[6] || ''
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

    for (const sheetName of VALID_SHEETS) {
        const ws = wb.addWorksheet(sheetName);
        const rows = data.sheets[sheetName] || [];

        // Row 1: Date de mise à jour globale uniquement
        ws.mergeCells('A1:G1');
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

        // Data rows
        rows.forEach((r, idx) => {
            const rowNum = idx + 3;
            const values = [
                r.codeBarre || '',
                r.codeArticle || '',
                r.designation || '',
                r.prixNormale || '',
                r.prixPromo || '',
                r.quantite || '',
                r.dlv || ''
            ];
            values.forEach((v, c) => {
                const cell = ws.getCell(rowNum, c + 1);
                cell.value = v;
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.border = thinBorder;
            });
            ws.getRow(rowNum).height = 18;
        });

        const widths = [16, 14, 45, 14, 12, 10, 12];
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
