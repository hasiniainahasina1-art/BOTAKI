export default async function handler(req, res) {
    // Accepter uniquement les requêtes POST
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Méthode non autorisée. Utilisez POST.' });
    }

    try {
        const { fileBase64, fileName, category } = req.body;

        // Vérifier que tous les champs requis sont présents
        if (!fileBase64 || !fileName || !category) {
            return res.status(400).json({ error: 'Paramètres manquants : fileBase64, fileName, category' });
        }

        // Récupérer le token GitHub stocké dans les variables d'environnement Vercel
        const token = process.env.GITHUB_TOKEN;
        if (!token) {
            throw new Error('Token GitHub non configuré sur Vercel (GITHUB_TOKEN)');
        }

        // Chemin complet du fichier dans le dépôt
        const path = `inventaires/${category}/${fileName}`;

        // Convertir le base64 en contenu utilisable par l'API GitHub
        const content = Buffer.from(fileBase64, 'base64').toString('base64');

        // Appeler l'API GitHub pour créer (ou mettre à jour) le fichier
        const response = await fetch(`https://api.github.com/repos/hasiniainahasina1-art/BOTAKI/contents/${path}`, {
            method: 'PUT',
            headers: {
                Authorization: `token ${token}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                message: `Ajout inventaire ${fileName}`,
                content: content,
                branch: 'main', // ou 'master' selon ta branche par défaut
            }),
        });

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.message || 'Erreur lors de l\'envoi vers GitHub');
        }

        const result = await response.json();
        res.status(200).json({ success: true, url: result.content.html_url });
    } catch (error) {
        console.error('Erreur upload-inventaire:', error);
        res.status(500).json({ error: error.message });
    }
}
