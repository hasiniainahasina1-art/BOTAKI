export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).json({ error: 'Méthode non autorisée' });

    const { password } = req.body;
    const adminPassword = process.env.ADMIN_PASSWORD;

    if (!adminPassword) {
        return res.status(500).json({ error: 'Mot de passe admin non configuré sur le serveur.' });
    }

    if (password === adminPassword) {
        return res.status(200).json({ success: true });
    } else {
        return res.status(401).json({ success: false });
    }
}
