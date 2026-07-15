import ExcelJS from 'exceljs';

export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Méthode non autorisée' });

    const { fileBase64, fileName, mode } = req.body;
    if (!fileBase64 || !fileName || !mode) {
        return res.status(400).json({ error: 'Paramètres manquants' });
    }

    const token = process.env.GITHUB_TOKEN;
    if (!token) {
        return res.status(500).json({ error: 'Token GitHub non configuré' });
    }

    const repoOwner = 'hasiniainahasina1-art';
    const repoName = 'BOTAKI';
    const filePath = fileName;

    try {
        const newBuffer = Buffer.from(fileBase64, 'base64');

        if (mode === 'replace') {
            const content = newBuffer.toString('base64');
            await commitFile(token, repoOwner, repoName, filePath, content, `Mise à jour ${fileName}`);
            return res.status(200).json({ success: true });
        }
        else if (mode === 'append') {
            const existingBuffer = await getFileContent(token, repoOwner, repoName, filePath);
            if (!existingBuffer) {
                return res.status(404).json({ error: 'Fichier existant introuvable' });
            }

            // Lire les deux fichiers Excel avec ExcelJS
            const oldWorkbook = new ExcelJS.Workbook();
            await oldWorkbook.xlsx.load(existingBuffer);
            const newWorkbook = new ExcelJS.Workbook();
            await newWorkbook.xlsx.load(newBuffer);

            const oldSheet = oldWorkbook.worksheets[0];
            const newSheet = newWorkbook.worksheets[0];

            // Récupérer les lignes sous forme de tableaux (valeurs)
            const oldRows = [];
            oldSheet.eachRow({ includeEmpty: true }, (row) => {
                oldRows.push(row.values.slice(1)); // row.values commence par undefined à l'index 0
            });
            const newRows = [];
            newSheet.eachRow({ includeEmpty: true }, (row) => {
                newRows.push(row.values.slice(1));
            });

            // Déterminer si l'ancien fichier a un en-tête (première ligne contenant des chaînes comme 'code')
            const hasHeader = oldRows.length > 0 && oldRows[0].some(cell => cell && cell.toString().toLowerCase().includes('code'));
            let header = hasHeader ? oldRows[0] : [];
            let oldBody = hasHeader ? oldRows.slice(1) : oldRows;

            // Pour le nouveau, on considère qu'il a la même structure. S'il a un en-tête, on l'ignore.
            let newBody = newRows;
            if (hasHeader && newRows.length > 0 && newRows[0].some(cell => cell && cell.toString().toLowerCase().includes('code'))) {
                newBody = newRows.slice(1);
            }

            // Fusionner
            const mergedBody = [...oldBody, ...newBody];
            const mergedData = header.length > 0 ? [header, ...mergedBody] : mergedBody;

            // Créer un nouveau fichier Excel
            const mergedWorkbook = new ExcelJS.Workbook();
            const ws = mergedWorkbook.addWorksheet(oldSheet.name);
            mergedData.forEach(row => {
                ws.addRow(row);
            });

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

async function getFileContent(token, owner, repo, path) {
    const url = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    const resp = await fetch(url, {
        headers: { Authorization: `token ${token}` }
    });
    if (!resp.ok) return null;
    const data = await resp.json();
    return Buffer.from(data.content, 'base64');
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

async function commitFile(token, owner, repo, path, contentBase64, message) {
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
