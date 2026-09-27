/**
 * API Vercel – Nettoyage automatique des inventaires de plus de 3 jours
 * Déclenchement : cron Vercel (voir vercel.json)
 * Méthode : GET (appelé par Vercel Cron) ou POST (appel manuel admin)
 */

const REPO_OWNER = 'hasiniainahasina1-art';
const REPO_NAME  = 'BOTAKI';
const CATEGORIES = ['Epicerie', 'BAZARD', 'Boisson'];
const DELAI_JOURS = 3;

/**
 * Extrait la date du nom de fichier.
 * Format attendu : inventaire_<code>_<YYYY-MM-DD>.xlsx
 * Ex : inventaire_KAC01_2026-07-27.xlsx → 2026-07-27
 */
function extraireDateFichier(fileName) {
    const match = fileName.match(/(\d{4}-\d{2}-\d{2})\.xlsx$/);
    if (!match) return null;
    const date = new Date(match[1]);
    return isNaN(date.getTime()) ? null : date;
}

/**
 * Récupère la liste des fichiers d'un dossier GitHub
 */
async function listerFichiers(token, category) {
    const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/inventaires/${category}`;
    const resp = await fetch(url, {
        headers: { Authorization: `token ${token}`, Accept: 'application/vnd.github.v3+json' }
    });
    if (!resp.ok) {
        console.error(`Erreur listing ${category}: ${resp.status}`);
        return [];
    }
    const data = await resp.json();
    return Array.isArray(data) ? data.filter(f => f.name.endsWith('.xlsx')) : [];
}

/**
 * Supprime un fichier sur GitHub
 */
async function supprimerFichier(token, filePath, sha) {
    const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${filePath}`;
    const resp = await fetch(url, {
        method: 'DELETE',
        headers: {
            Authorization: `token ${token}`,
            'Content-Type': 'application/json',
            Accept: 'application/vnd.github.v3+json'
        },
        body: JSON.stringify({
            message: `Suppression automatique (expirés > ${DELAI_JOURS}j) : ${filePath}`,
            sha: sha,
            branch: 'main'
        })
    });
    return resp.ok;
}

export default async function handler(req, res) {
    // Accepter GET (cron Vercel) et POST (appel admin manuel)
    if (req.method !== 'GET' && req.method !== 'POST') {
        return res.status(405).json({ error: 'Méthode non autorisée' });
    }

    const token = process.env.GITHUB_TOKEN;
    if (!token) {
        return res.status(500).json({ error: 'Token GitHub non configuré (GITHUB_TOKEN)' });
    }

    const maintenant = new Date();
    const limitDate  = new Date(maintenant.getTime() - DELAI_JOURS * 24 * 60 * 60 * 1000);

    const rapport = { supprimes: [], ignores: [], erreurs: [] };

    for (const category of CATEGORIES) {
        const fichiers = await listerFichiers(token, category);

        for (const fichier of fichiers) {
            // Ignorer les fichiers sentinelles (.gitkeep, etc.)
            if (!fichier.name.endsWith('.xlsx')) continue;

            const dateFichier = extraireDateFichier(fichier.name);

            if (!dateFichier) {
                rapport.ignores.push(`${category}/${fichier.name} (date introuvable dans le nom)`);
                continue;
            }

            if (dateFichier < limitDate) {
                // Fichier expiré → suppression
                const ok = await supprimerFichier(token, fichier.path, fichier.sha);
                if (ok) {
                    rapport.supprimes.push(`${category}/${fichier.name}`);
                } else {
                    rapport.erreurs.push(`${category}/${fichier.name}`);
                }
            } else {
                rapport.ignores.push(`${category}/${fichier.name} (encore valide)`);
            }
        }
    }

    console.log('[cleanup-inventaires]', JSON.stringify(rapport));

    return res.status(200).json({
        success: true,
        date_execution: maintenant.toISOString(),
        delai_jours: DELAI_JOURS,
        ...rapport
    });
}
