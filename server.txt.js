require('dotenv').config();

// Correctif (Polyfill) pour rendre 'fetch' disponible globalement
if (!globalThis.fetch) {
  const fetch = require('node-fetch');
  globalThis.fetch = fetch;
  globalThis.Headers = fetch.Headers;
  globalThis.Request = fetch.Request;
  globalThis.Response = fetch.Response;
}

const express = require('express');
const mongoose = require('mongoose');
const http = require('http');
const path = require('path');
const cors = require('cors');
const fileUpload = require('express-fileupload');
const XLSX = require('xlsx');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { GoogleGenerativeAI } = require("@google/generative-ai");

const app = express();
const server = http.createServer(app);

const PORT = process.env.PORT || 3000;
const WORD_TEMPLATE_URL = process.env.WORD_TEMPLATE_URL;

// Initialisation de l'API Gemini
let geminiModel;
if (process.env.GEMINI_API_KEY) {
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    // Utilisation d'un modèle plus récent et moins sujet aux limites de quota du free tier
    geminiModel = genAI.getGenerativeModel({ model: "gemini-1.5-flash-latest" });
    console.log('✅ SDK Google Gemini initialisé avec le modèle gemini-1.5-flash-latest.');
} else {
    console.warn('⚠️ GEMINI_API_KEY non défini dans le fichier .env. La génération de plans de leçon par IA sera désactivée.');
}

// Dates spécifiques à chaque semaine (côté serveur)
const specificWeekDateRangesNode = {
     1: { start: '2024-08-25', end: '2024-08-29' },  2: { start: '2024-09-01', end: '2024-09-05' },
     3: { start: '2024-09-08', end: '2024-09-12' },  4: { start: '2024-09-15', end: '2024-09-19' },
     5: { start: '2024-09-22', end: '2024-09-26' },  6: { start: '2024-09-29', end: '2024-10-03' },
     7: { start: '2024-10-06', end: '2024-10-10' },  8: { start: '2024-10-13', end: '2024-10-17' },
     9: { start: '2024-10-20', end: '2024-10-24' }, 10: { start: '2024-10-27', end: '2024-10-31' },
    11: { start: '2024-11-03', end: '2024-11-07' }, 12: { start: '2024-11-10', end: '2024-11-14' },
    13: { start: '2024-11-17', end: '2024-11-21' }, 14: { start: '2024-11-24', end: '2024-11-28' },
    15: { start: '2024-12-01', end: '2024-12-05' }, 16: { start: '2024-12-08', end: '2024-12-12' },
    17: { start: '2024-12-15', end: '2024-12-19' }, 18: { start: '2024-12-22', end: '2024-12-26' },
    19: { start: '2024-12-29', end: '2025-01-02' }, 20: { start: '2025-01-05', end: '2025-01-09' },
    21: { start: '2025-01-12', end: '2025-01-16' }, 22: { start: '2025-01-19', end: '2025-01-23' },
    23: { start: '2025-01-26', end: '2025-01-30' }, 24: { start: '2025-02-02', end: '2025-02-06' },
    25: { start: '2025-02-09', end: '2025-02-13' }, 26: { start: '2025-02-16', end: '2025-02-20' },
    27: { start: '2025-02-23', end: '2025-02-27' }, 28: { start: '2025-03-02', end: '2025-03-06' },
    29: { start: '2025-03-09', end: '2025-03-13' }, 30: { start: '2025-03-16', end: '2025-03-20' },
    31: { start: '2025-03-23', end: '2025-03-27' }, 32: { start: '2025-03-30', end: '2025-04-03' },
    33: { start: '2025-04-06', end: '2025-04-10' }, 34: { start: '2025-04-13', end: '2025-04-17' },
    35: { start: '2025-04-20', end: '2025-04-24' }, 36: { start: '2025-04-27', end: '2025-05-01' },
    37: { start: '2025-05-04', end: '2025-05-08' }, 38: { start: '2025-05-11', end: '2025-05-15' },
    39: { start: '2025-05-18', end: '2025-05-22' }, 40: { start: '2025-05-25', end: '2025-05-29' },
    41: { start: '2025-06-01', end: '2025-06-05' }, 42: { start: '2025-06-08', end: '2025-06-12' },
    43: { start: '2025-06-15', end: '2025-06-19' }, 44: { start: '2025-06-22', end: '2025-06-26' },
    45: { start: '2025-06-29', end: '2025-07-03' }, 46: { start: '2025-07-06', end: '2025-07-10' },
    47: { start: '2025-07-13', end: '2025-07-17' }, 48: { start: '2025-07-20', end: '2025-07-24' }
};

// --- Middleware ---
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(__dirname));
app.use(fileUpload());

// --- Connexion MongoDB ---
const MONGO_URL = process.env.MONGO_URL;
if (!MONGO_URL) { console.error('❌ FATAL: MONGO_URL non définie.'); process.exit(1); }
mongoose.connect(MONGO_URL, { useNewUrlParser: true, useUnifiedTopology: true, useFindAndModify: false })
    .then(() => console.log('✅ Connecté à MongoDB'))
    .catch(err => { console.error('❌ FATAL: Erreur connexion MongoDB:', err); process.exit(1); });
mongoose.connection.on('error', err => { console.error('❌ Erreur MongoDB post-connexion:', err); });
mongoose.connection.on('disconnected', () => { console.warn('⚠️ Déconnecté de MongoDB'); });

// --- Schéma et Modèle Mongoose ---
const PlanSchema = new mongoose.Schema({
    week: { type: Number, required: true, index: true, unique: true },
    data: { type: Array, required: true },
    classNotes: { type: Map, of: String, default: {} }
}, { timestamps: true });
PlanSchema.index({ week: 1 });
const Plan = mongoose.model('Plan', PlanSchema);

// --- Utilisateurs Valides ---
const validUsers = {
    "Zine": "Zine", "Abas": "Abas", "Tonga": "Tonga", "Ilyas": "Ilyas", "Morched": "Morched",
    "عبد الرحمان": "عبد الرحمان", "Youssif": "Youssif", "عبد العزيز": "عبد العزيز",
    "Med Ali": "Med Ali", "Sami": "Sami", "جابر": "جابر", "محمد الزبيدي": "محمد الزبيدي",
    "فارس": "فارس", "AutreProf": "AutreProf", "Mohamed": "Mohamed"
};
console.log(`Utilisateurs login: ${Object.keys(validUsers).join(', ')}`);

// --- Fonctions utilitaires dates ---
function formatDateFrenchNode(date) { if (!date || isNaN(date.getTime())) { return "Date invalide"; } const days = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]; const months = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]; const dayName = days[date.getUTCDay()]; const dayNum = String(date.getUTCDate()).padStart(2, '0'); const monthName = months[date.getUTCMonth()]; const yearNum = date.getUTCFullYear(); return `${dayName} ${dayNum} ${monthName} ${yearNum}`; }
function getDateForDayNameNode(weekStartDate, dayName) { if (!weekStartDate || isNaN(weekStartDate.getTime())) return null; const dayOrder = { "Dimanche": 0, "Lundi": 1, "Mardi": 2, "Mercredi": 3, "Jeudi": 4 }; const offset = dayOrder[dayName]; if (offset === undefined) return null; const specificDate = new Date(Date.UTC(weekStartDate.getUTCFullYear(), weekStartDate.getUTCMonth(), weekStartDate.getUTCDate())); specificDate.setUTCDate(specificDate.getUTCDate() + offset); return specificDate; }

// --- Routes ---

app.get('/', (req, res) => { res.sendFile(path.join(__dirname, 'index.html')); });

app.post('/login', (req, res) => {
    const { username, password } = req.body;
    console.log(`Tentative connexion : Utilisateur='${username}'`);
    if (validUsers[username] && validUsers[username] === password) {
        console.log(`Connexion réussie pour : ${username}`);
        res.status(200).json({ success: true, username: username });
    } else {
        console.log(`Échec connexion pour : ${username}`);
        res.status(401).json({ success: false, message: 'Identifiants invalides' });
    }
});
app.post('/save-plan', async (req, res) => {
    console.log('--- Requête /save-plan reçue ---');
    try {
        const { week, data } = req.body; const weekNumber = Number(week);
        if (!Number.isInteger(weekNumber) || weekNumber <= 0 || weekNumber > 53) return res.status(400).json({ message: 'Semaine invalide.' });
        if (!Array.isArray(data)) return res.status(400).json({ message: '"data" doit être un tableau.' });
        console.log(`[SAVE-PLAN] Sauvegarde S${weekNumber}, Lignes: ${data.length}`);
        const cleanedData = data.map(item => { if (item && typeof item === 'object' && !Array.isArray(item)) { const newItem = { ...item }; delete newItem._id; delete newItem.id; if (newItem.hasOwnProperty('updatedAt') && newItem.updatedAt && isNaN(new Date(newItem.updatedAt).getTime())) { delete newItem.updatedAt; } return newItem; } return null; }).filter(Boolean);
        let updatedPlan; try { updatedPlan = await Plan.findOneAndUpdate( { week: weekNumber }, { $set: { data: cleanedData } }, { new: true, upsert: true, runValidators: true, setDefaultsOnInsert: true } ); } catch (dbError) { console.error(`❌ Erreur DB /save-plan S${weekNumber}:`, dbError); if (dbError.code === 11000) return res.status(500).json({ message: `Erreur E11000. Semaine ${weekNumber} existe déjà?` }); return res.status(500).json({ message: `Erreur DB sauvegarde plan: ${dbError.message}` }); }
        console.log(`[SAVE-PLAN] Données S${weekNumber} OK. Doc ID: ${updatedPlan?._id}`);
        res.status(200).json({ message: `Tableau S${weekNumber} enregistré.` });
    } catch (error) { console.error(`❌ Erreur serveur /save-plan S${req.body?.week}:`, error); res.status(500).json({ message: `Erreur interne /save-plan: ${error.message}`, error: process.env.NODE_ENV !== 'production' ? error.toString() : undefined }); }
    console.log('--- Fin /save-plan ---');
});
app.post('/save-notes', async (req, res) => {
    console.log('--- Requête /save-notes reçue ---');
    try {
        const { week, classe, notes } = req.body; const weekNumber = Number(week);
        if (!Number.isInteger(weekNumber) || weekNumber <= 0 || weekNumber > 53) { return res.status(400).json({ message: 'Semaine invalide.' }); }
        if (!classe || typeof classe !== 'string' || classe.trim() === '') { return res.status(400).json({ message: 'Classe invalide.' }); }
        if (notes == null || typeof notes !== 'string') { return res.status(400).json({ message: 'Notes invalides (doit être string).' }); }
        const classeKeyForMap = classe.trim();
        console.log(`[SAVE-NOTES] Demande S${weekNumber}, Classe:"${classeKeyForMap}", Longueur:${notes.length}`);
        let result;
        try {
            const updateOperation = { $set: { [`classNotes.${classeKeyForMap}`]: notes } };
            console.log('[SAVE-NOTES] Op MongoDB:', JSON.stringify(updateOperation));
            result = await Plan.updateOne( { week: weekNumber }, updateOperation, { upsert: true } );
        } catch (dbError) { console.error(`❌ Erreur DB /save-notes S${weekNumber}, C:"${classeKeyForMap}":`, dbError); return res.status(500).json({ message: `Erreur DB /save-notes: ${dbError.message}` }); }
        console.log(`[SAVE-NOTES] Résultat DB S${weekNumber}, C:"${classeKeyForMap}":`, result);
        if (result.acknowledged || (result.ok === 1)) {
             if (result.upsertedCount > 0 || (result.upserted && result.upserted.length > 0)) console.log(`[SAVE-NOTES] Doc créé et note ajoutée S${weekNumber}, C:${classeKeyForMap}`);
             else if (result.modifiedCount > 0 || result.nModified > 0) console.log(`[SAVE-NOTES] Note MAJ S${weekNumber}, C:${classeKeyForMap}`);
             else if (result.matchedCount > 0 || result.n > 0) console.log(`[SAVE-NOTES] Note identique S${weekNumber}, C:${classeKeyForMap}`);
             else console.warn(`[SAVE-NOTES] Op DB confirmée mais rien MAJ/créé? S${weekNumber}, C:${classeKeyForMap}`, result);
             res.status(200).json({ message: `Note pour ${classeKeyForMap} (S${weekNumber}) enregistrée.` });
        } else { console.error(`[SAVE-NOTES] Op sauvegarde note non confirmée DB S${weekNumber}, C:${classeKeyForMap}`, result); throw new Error("Sauvegarde non confirmée."); }
    } catch (error) { console.error(`❌ Erreur serveur /save-notes S${req.body?.week}, C:${req.body?.classe}:`, error); res.status(500).json({ message: `Erreur interne /save-notes: ${error.message}`, error: process.env.NODE_ENV !== 'production' ? error.toString() : undefined }); }
    console.log('--- Fin /save-notes ---');
});
app.post('/save-row', async (req, res) => {
    console.log('\n--- Requête /save-row reçue ---');
    try {
        const { week, data: rowData } = req.body; const weekNumber = Number(week);
        if (!Number.isInteger(weekNumber) || weekNumber <= 0 || weekNumber > 53) return res.status(400).json({ message: 'Semaine invalide.' });
        if (!rowData || typeof rowData !== 'object' || Array.isArray(rowData) || Object.keys(rowData).length === 0) return res.status(400).json({ message: 'Données ligne invalides.' });
        console.log(`[SAVE-ROW] Demande S${weekNumber}, Début:`, JSON.stringify(rowData).substring(0, 150) + '...');
        const findKey = (target) => Object.keys(rowData).find(k => k.trim().toLowerCase() === target.toLowerCase());
        const ensKey = findKey('Enseignant'), clsKey = findKey('Classe'), jourKey = findKey('Jour'), perKey = findKey('Période'), matKey = findKey('Matière');
        const uniqueKeyFieldsMatch = {}; const requiredKeys = { 'Enseignant': ensKey, 'Classe': clsKey, 'Jour': jourKey, 'Période': perKey, 'Matière': matKey };
        for (const [name, key] of Object.entries(requiredKeys)) { if (!key || rowData[key] == null || String(rowData[key]).trim() === '') { console.error(`[SAVE-ROW] Clé '${name}' ('${key}') manquante/vide.`); return res.status(400).json({ message: `Champ clé '${name}' manquant/vide.` }); } uniqueKeyFieldsMatch[key] = rowData[key]; }
        const rootQuery = { week: weekNumber, data: { $elemMatch: uniqueKeyFieldsMatch } };
        const lessonKey = findKey('Leçon'), classWorkKey = findKey('Travaux de classe'), supportKey = findKey('Support'), homeworkKey = findKey('Devoirs'), updatedAtKey = findKey('updatedAt');
        const updateFields = {}; const now = new Date();
        if (lessonKey && rowData.hasOwnProperty(lessonKey)) updateFields[`data.$.${lessonKey}`] = rowData[lessonKey];
        if (classWorkKey && rowData.hasOwnProperty(classWorkKey)) updateFields[`data.$.${classWorkKey}`] = rowData[classWorkKey];
        if (supportKey && rowData.hasOwnProperty(supportKey)) updateFields[`data.$.${supportKey}`] = rowData[supportKey];
        if (homeworkKey && rowData.hasOwnProperty(homeworkKey)) updateFields[`data.$.${homeworkKey}`] = rowData[homeworkKey];
        const finalUpdatedAtKeyName = updatedAtKey || 'updatedAt'; updateFields[`data.$.${finalUpdatedAtKeyName}`] = now;
        const updateOperation = { $set: updateFields }; console.log('[SAVE-ROW] Query:', JSON.stringify(rootQuery)); console.log('[SAVE-ROW] Update Op:', JSON.stringify(updateOperation));
        let result; try { result = await Plan.updateOne(rootQuery, updateOperation); } catch (dbError) { console.error(`❌ Erreur DB /save-row S${weekNumber}:`, dbError); return res.status(500).json({ message: `Erreur DB /save-row: ${dbError.message}` }); }
        console.log('[SAVE-ROW] Résultat updateOne:', result);
        if (result.n === 0 && result.nModified === 0) {
            console.error(`[SAVE-ROW] Ligne non trouvée MAJ S${weekNumber} query:`, rootQuery);
            return res.status(404).json({ message: 'Ligne non trouvée pour la mise à jour. Les champs clés ont-ils été modifiés par erreur ailleurs ?' });
        }
        if (result.nModified >= 0 || result.n > 0) {
            const updatedDataObject = { [finalUpdatedAtKeyName]: now };
            console.log(`[SAVE-ROW] Ligne enregistrée/traitée OK S${weekNumber}`);
            return res.status(200).json({ message: 'Ligne enregistrée.', updatedData: updatedDataObject });
        } else { console.error(`[SAVE-ROW] Résultat inattendu updateOne S${weekNumber}:`, result); return res.status(500).json({ message: 'Résultat inattendu mise à jour.' }); }
    } catch (error) {
        console.error(`❌ Erreur serveur /save-row S${req.body?.week}:`, error);
        if (!res.headersSent) {
             res.status(500).json({
                message: `Erreur interne /save-row: ${error.message}`,
                error: process.env.NODE_ENV !== 'production' ? error.toString() : undefined
            });
        }
    }
    console.log('--- Fin /save-row ---');
});
app.get('/plans/:week', async (req, res) => {
    const requestedWeek = req.params.week; console.log(`--- Requête /plans/${requestedWeek} ---`);
    try { const weekNumber = parseInt(requestedWeek, 10); if (isNaN(weekNumber) || weekNumber <= 0 || weekNumber > 53) return res.status(400).json({ message: 'Semaine invalide.' });
        let planDocument; try { planDocument = await Plan.findOne({ week: weekNumber }, 'data classNotes').lean(); } catch (dbError) { console.error(`❌ Erreur DB /plans/${requestedWeek}:`, dbError); return res.status(500).json({ message: 'Erreur DB récupération plan.' }); }
        if (!planDocument) { console.log(`[GET /plans] Doc non trouvé S${weekNumber}.`); return res.status(200).json({ planData: [], classNotes: {} }); }
        const notesToSend = planDocument.classNotes instanceof Map ? Object.fromEntries(planDocument.classNotes) : (planDocument.classNotes || {});
        console.log(`[GET /plans] Doc trouvé S${weekNumber}. Lignes:${planDocument.data?.length || 0}. Notes:${Object.keys(notesToSend).length}`);
        res.status(200).json({ planData: planDocument.data || [], classNotes: notesToSend });
    } catch (error) { console.error(`❌ Erreur serveur /plans/${requestedWeek}:`, error); res.status(500).json({ message: 'Erreur interne /plans.', error: process.env.NODE_ENV !== 'production' ? error.toString() : undefined }); }
    console.log(`--- Fin /plans/${requestedWeek} ---`);
});
app.post('/generate-word', async (req, res) => {
    console.log('--- Requête /generate-word reçue ---');
    try { const { week, classe, data, notes } = req.body; const weekNumber = Number(week); if (!Number.isInteger(weekNumber) || weekNumber <= 0 || weekNumber > 53) return res.status(400).json({ message: 'Semaine invalide.' }); if (!classe || typeof classe !== 'string') return res.status(400).json({ message: 'Classe invalide.' }); if (!Array.isArray(data)) return res.status(400).json({ message: '"data" doit être array.' }); const finalNotes = (typeof notes === 'string') ? notes : ""; console.log(`[GEN-WORD] Demande S${weekNumber}, C:'${classe}', Lignes:${data.length}, Note:${finalNotes ? 'Oui' : 'Non'}`);
        let templateBuffer; try { console.log(`[GEN-WORD] Téléchargement modèle...`); const response = await fetch(WORD_TEMPLATE_URL); if (!response.ok) throw new Error(`Échec modèle Word (${response.status})`); templateBuffer = Buffer.from(await response.arrayBuffer()); console.log(`[GEN-WORD] Modèle OK (${templateBuffer.length} o).`); } catch (e) { console.error(`[GEN-WORD] ERREUR modèle:`, e); return res.status(500).json({ message: `Erreur récup modèle Word.` }); }
        const zip = new PizZip(templateBuffer); let doc; try { doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => "" }); } catch (e) { console.error(`[GEN-WORD] Erreur init Docxtemplater:`, e); return res.status(500).json({ message: 'Erreur init générateur.' }); }
        console.log("[GEN-WORD] Transformation data..."); const groupedByDay = {}; const dayOrder = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]; const datesNode = specificWeekDateRangesNode[weekNumber];
        let weekStartDateNode = null; if (datesNode?.start) { try { weekStartDateNode = new Date(datesNode.start + 'T00:00:00Z'); if (isNaN(weekStartDateNode.getTime())) throw new Error('Date Invalide'); } catch (e) { weekStartDateNode = null; console.error(`[GEN-WORD] Date début invalide S${weekNumber}: ${datesNode.start}`); } }
        if (!weekStartDateNode) { console.error(`[GEN-WORD] Dates serveur MANQUANTES S${weekNumber}. Annulation.`); return res.status(500).json({ message: `Config Erreur: Dates serveur manquantes pour S${weekNumber}.` }); }
        const sampleRow = data[0] || {}; const findHeaderKey = (target) => Object.keys(sampleRow).find(k => k?.trim().toLowerCase() === target.toLowerCase()) || target; const jourKey = findHeaderKey('Jour'), periodeKey = findHeaderKey('Période'), matiereKey = findHeaderKey('Matière'), leconKey = findHeaderKey('Leçon'), travauxKey = findHeaderKey('Travaux de classe'), supportKey = findHeaderKey('Support'), devoirsKey = findHeaderKey('Devoirs');
        data.forEach(item => { if (!item || typeof item !== 'object') return; const day = item[jourKey]; if (day && dayOrder.includes(day)) { if (!groupedByDay[day]) groupedByDay[day] = []; groupedByDay[day].push(item); } });
        const joursData = dayOrder.map(dayName => { if (groupedByDay[dayName]) { const dateOfDay = getDateForDayNameNode(weekStartDateNode, dayName); const formattedDate = dateOfDay ? formatDateFrenchNode(dateOfDay) : dayName; const sortedEntries = groupedByDay[dayName].sort((a, b) => { const pA = parseInt(a[periodeKey], 10), pB = parseInt(b[periodeKey], 10); if (!isNaN(pA) && !isNaN(pB)) return pA - pB; return String(a[periodeKey] ?? "").localeCompare(String(b[periodeKey] ?? "")); }); const matieres = sortedEntries.map(item => ({ matiere: item[matiereKey] ?? "", Lecon: item[leconKey] ?? "", travailDeClasse: item[travauxKey] ?? "", Support: item[supportKey] ?? "", devoirs: item[devoirsKey] ?? "" })); return { jourDateComplete: formattedDate, matieres: matieres }; } return null; }).filter(Boolean);
        let plageSemaineText = `Semaine ${weekNumber}`; if (datesNode?.start && datesNode?.end) { try { const startD = new Date(datesNode.start + 'T00:00:00Z'), endD = new Date(datesNode.end + 'T00:00:00Z'); if (!isNaN(startD.getTime()) && !isNaN(endD.getTime())) { const startS = formatDateFrenchNode(startD).replace(/^./, c => c.toUpperCase()).replace(/ (\d{2}) /, ' le $1 '); const endS = formatDateFrenchNode(endD).replace(/^./, c => c.toUpperCase()); plageSemaineText = `du ${startS} à ${endS}`; } } catch (e) { console.error("[GEN-WORD] Erreur formatage plage:", e); } }
        const templateData = { semaine: weekNumber, classe: classe, jours: joursData, notes: finalNotes, plageSemaine: plageSemaineText }; console.log("[GEN-WORD] Rendu doc..."); try { doc.render(templateData); } catch (error) { console.error('[GEN-WORD] Erreur rendu:', error); if (error.properties?.errors) { const dErrors = error.properties.errors.map(err => `[Tag:${err.id}] ${err.message}`).join('; '); console.error('[GEN-WORD] Erreurs template:', dErrors); return res.status(500).json({ message: `Erreur template: ${error.message}. Détails: ${dErrors}`, error: error.toString() }); } return res.status(500).json({ message: `Erreur rendu: ${error.message}`, error: error.toString() }); }
        console.log("[GEN-WORD] Génération buffer..."); const buf = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        const safeClasseName = classe.replace(/[^a-z0-9]/gi, '_').replace(/_+/g, '_'); const filename = `Plan_hebdomadaire_S${weekNumber}_${safeClasseName}.docx`; console.log(`[GEN-WORD] Envoi: ${filename} (${buf.length} o)`); res.setHeader('Content-Disposition', `attachment; filename="${filename}"`); res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'); res.send(buf);
    } catch (error) { console.error('❌ Erreur serveur majeure /generate-word:', error); if (!res.headersSent) res.status(500).json({ message: 'Erreur interne /generate-word.', error: process.env.NODE_ENV !== 'production' ? error.toString() : undefined }); }
    console.log('--- Fin /generate-word ---');
});


// =========================================================================
// == DÉBUT DE LA SECTION MODIFIÉE : NOUVELLE LOGIQUE POUR L'EXPORT EXCEL ==
// =========================================================================
app.post('/generate-excel-workbook', async (req, res) => {
    console.log('--- Requête /generate-excel-workbook (logique feuille unique) ---');
    try {
        const { week } = req.body;
        const weekNumber = Number(week);
        if (!Number.isInteger(weekNumber) || weekNumber <= 0 || weekNumber > 53) {
            return res.status(400).json({ message: 'Semaine invalide.' });
        }
        console.log(`[GEN-EXCEL-SINGLE] Préparation du fichier pour la S${weekNumber}`);

        // 1. Récupérer les données de la base de données
        let planDocument;
        try {
            planDocument = await Plan.findOne({ week: weekNumber }, 'data').lean();
        } catch (dbError) {
            console.error(`❌ Erreur DB /generate-excel-workbook S${weekNumber}:`, dbError);
            return res.status(500).json({ message: 'Erreur DB récupération des données.' });
        }

        if (!planDocument?.data?.length) {
            console.log(`[GEN-EXCEL-SINGLE] Aucune donnée trouvée en DB pour la S${weekNumber}.`);
            return res.status(404).json({ message: `Aucune donnée trouvée pour la semaine ${weekNumber}.` });
        }

        const allData = planDocument.data;
        console.log(`[GEN-EXCEL-SINGLE] ${allData.length} lignes récupérées depuis la DB pour la S${weekNumber}.`);

        // 2. Définir les en-têtes et l'ordre des colonnes demandés
        const finalHeaders = [
            'Enseignant',
            'Jour',
            'Période',
            'Classe',
            'Matière',
            'Leçon',
            'Travaux de classe',
            'Support',
            'Devoirs'
        ];
        
        // Fonction utilitaire pour trouver la clé réelle (insensible à la casse) dans un objet de données
        const findKey = (item, targetHeader) => {
            if (!item || typeof item !== 'object') return undefined;
            const targetLower = targetHeader.toLowerCase().trim();
            return Object.keys(item).find(k => k.toLowerCase().trim() === targetLower);
        };

        // 3. Formater les données pour correspondre à la structure souhaitée
        const formattedData = allData.map(item => {
            const row = {};
            finalHeaders.forEach(header => {
                const itemKey = findKey(item, header);
                // Si la clé est trouvée dans l'objet, on prend sa valeur, sinon une chaîne vide
                row[header] = itemKey ? item[itemKey] : '';
            });
            return row;
        });

        console.log(`[GEN-EXCEL-SINGLE] ${formattedData.length} lignes formatées pour l'export.`);

        // 4. Créer la feuille de calcul unique
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(formattedData, { header: finalHeaders });

        // 5. (Optionnel mais recommandé) Ajuster la largeur des colonnes pour une meilleure lisibilité
        worksheet['!cols'] = [
            { wch: 20 }, // Enseignant
            { wch: 15 }, // Jour
            { wch: 10 }, // Période
            { wch: 12 }, // Classe
            { wch: 20 }, // Matière
            { wch: 45 }, // Leçon
            { wch: 45 }, // Travaux de classe
            { wch: 25 }, // Support
            { wch: 45 }  // Devoirs
        ];

        // 6. Ajouter la feuille au classeur
        const sheetName = `Plan S${weekNumber}`;
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

        // 7. Générer le buffer et envoyer le fichier au client
        console.log("[GEN-EXCEL-SINGLE] Génération du buffer...");
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const filename = `Plan_Hebdomadaire_S${weekNumber}_Complet.xlsx`;
        
        console.log(`[GEN-EXCEL-SINGLE] Envoi du fichier : ${filename} (${buffer.length} octets)`);
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (error) {
        console.error('❌ Erreur serveur majeure /generate-excel-workbook:', error);
        if (!res.headersSent) {
            res.status(500).json({ message: 'Erreur interne lors de la génération du fichier Excel.', error: process.env.NODE_ENV !== 'production' ? error.toString() : undefined });
        }
    }
    console.log('--- Fin /generate-excel-workbook ---');
});
// =========================================================================
// == FIN DE LA SECTION MODIFIÉE                                          ==
// =========================================================================


// ===== NOUVEAU ENDPOINT : Fournir la liste de toutes les classes uniques de la DB =====
app.get('/api/all-classes', async (req, res) => {
    console.log('--- Requête /api/all-classes reçue ---');
    try {
        // Utilise `distinct` pour obtenir un tableau de toutes les valeurs uniques pour le champ 'Classe'
        // dans le tableau 'data' de tous les documents.
        // On filtre les valeurs nulles ou vides.
        const classes = await Plan.distinct('data.Classe', { 'data.Classe': { $ne: null, $ne: "" } });
        console.log(`[ALL-CLASSES] ${classes.length} classes uniques trouvées.`);
        res.status(200).json(classes);
    } catch (error) {
        console.error('❌ Erreur serveur /api/all-classes:', error);
        res.status(500).json({ message: 'Erreur interne lors de la récupération des classes.' });
    }
    console.log('--- Fin /api/all-classes ---');
});


// ===== NOUVEAU ENDPOINT : Générer le rapport complet pour une classe donnée =====
app.post('/api/full-report-by-class', async (req, res) => {
    console.log('--- Requête /api/full-report-by-class reçue ---');
    try {
        const { classe: requestedClass } = req.body;
        if (!requestedClass) {
            return res.status(400).json({ message: 'Le nom de la classe est requis.' });
        }
        console.log(`[FULL-REPORT] Génération pour la classe: ${requestedClass}`);

        // 1. Récupérer toutes les données de toutes les semaines
        const allPlans = await Plan.find({}).sort({ week: 1 }).lean();
        if (!allPlans || allPlans.length === 0) {
            return res.status(404).json({ message: 'Aucune donnée trouvée dans la base de données.' });
        }
        console.log(`[FULL-REPORT] ${allPlans.length} semaines de données récupérées.`);

        // 2. Traiter les données
        const dataBySubject = {};
        const monthsFrench = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"];
        
        allPlans.forEach(plan => {
            const weekNumber = plan.week;
            const weekData = plan.data || [];

            // Déterminer le mois pour cette semaine
            let monthName = 'N/A';
            const weekDates = specificWeekDateRangesNode[weekNumber];
            if (weekDates && weekDates.start) {
                try {
                    const startDate = new Date(weekDates.start + 'T00:00:00Z');
                    monthName = monthsFrench[startDate.getUTCMonth()];
                } catch (e) {
                    console.warn(`[FULL-REPORT] Date invalide pour la semaine ${weekNumber}`);
                }
            }

            // Filtrer et organiser les données de la classe demandée
            weekData.forEach(item => {
                // Utilisation d'une recherche insensible à la casse pour les clés d'objet
                const findKey = (target) => Object.keys(item).find(k => k.trim().toLowerCase() === target.toLowerCase());

                const itemClassKey = findKey('classe');
                const itemSubjectKey = findKey('matière');
                
                if (itemClassKey && item[itemClassKey] === requestedClass && itemSubjectKey && item[itemSubjectKey]) {
                    const subject = item[itemSubjectKey];
                    if (!dataBySubject[subject]) {
                        dataBySubject[subject] = [];
                    }

                    // Créer l'objet ligne pour la feuille Excel
                    const row = {
                        'Mois': monthName,
                        'Semaine': weekNumber,
                        'Période': item[findKey('période')] || '',
                        'Leçon': item[findKey('leçon')] || '',
                        'Travaux de classe': item[findKey('travaux de classe')] || '',
                        'Support': item[findKey('support')] || '',
                        'Devoirs': item[findKey('devoirs')] || ''
                    };
                    dataBySubject[subject].push(row);
                }
            });
        });
        
        const subjectsFound = Object.keys(dataBySubject);
        if (subjectsFound.length === 0) {
            return res.status(404).json({ message: `Aucune donnée trouvée pour la classe '${requestedClass}'.` });
        }
        console.log(`[FULL-REPORT] Données organisées pour ${subjectsFound.length} matières.`);

        // 3. Générer le fichier Excel
        const workbook = XLSX.utils.book_new();
        const headers = ['Mois', 'Semaine', 'Période', 'Leçon', 'Travaux de classe', 'Support', 'Devoirs'];

        subjectsFound.sort().forEach(subject => {
            const safeSheetName = subject.substring(0, 30).replace(/[*?:/\\\[\]]/g, '_');
            const sheetData = dataBySubject[subject];
            
            console.log(`[FULL-REPORT] Création de la feuille '${safeSheetName}' avec ${sheetData.length} lignes.`);
            
            const worksheet = XLSX.utils.json_to_sheet(sheetData, { header: headers });

            // Ajuster la largeur des colonnes
            worksheet['!cols'] = [
                { wch: 12 }, // Mois
                { wch: 10 }, // Semaine
                { wch: 10 }, // Période
                { wch: 40 }, // Leçon
                { wch: 40 }, // Travaux de classe
                { wch: 25 }, // Support
                { wch: 40 }  // Devoirs
            ];

            XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);
        });

        // 4. Envoyer le fichier
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        const safeClassName = requestedClass.replace(/[^a-z0-9]/gi, '_');
        const filename = `Rapport_Complet_${safeClassName}.xlsx`;
        
        console.log(`[FULL-REPORT] Envoi du fichier: ${filename} (${buffer.length} octets)`);
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (error) {
        console.error('❌ Erreur serveur /api/full-report-by-class:', error);
        if (!res.headersSent) {
            res.status(500).json({ message: 'Erreur interne lors de la génération du rapport complet.' });
        }
    }
    console.log('--- Fin /api/full-report-by-class ---');
});


// ENDPOINT /generate-ai-lesson-plan (MODIFIÉ pour générer un WORD)
function getRowDataValue(rowData, keyName, defaultValue = "") {
    if (!rowData || typeof rowData !== 'object') return defaultValue;
    const actualKey = Object.keys(rowData).find(k => k.trim().toLowerCase() === keyName.trim().toLowerCase());
    return actualKey ? (rowData[actualKey] || defaultValue) : defaultValue;
}

app.post('/generate-ai-lesson-plan', async (req, res) => {
    console.log('--- Requête /generate-ai-lesson-plan reçue (Génération WORD) ---');
    if (!geminiModel) {
        return res.status(503).json({ message: "Service IA (Gemini) non configuré ou clé API manquante sur le serveur." });
    }

    try {
        const { week, rowData } = req.body;
        const weekNumber = Number(week);

        if (!Number.isInteger(weekNumber) || weekNumber <= 0 || weekNumber > 53) {
            return res.status(400).json({ message: 'Semaine invalide.' });
        }
        if (!rowData || typeof rowData !== 'object') {
            return res.status(400).json({ message: 'Données de ligne (rowData) invalides.' });
        }

        console.log(`[AI-PLAN] Demande S${weekNumber}, Data:`, JSON.stringify(rowData).substring(0, 200) + "...");

        const enseignant = getRowDataValue(rowData, 'Enseignant');
        const jour = getRowDataValue(rowData, 'Jour');
        const matiere = getRowDataValue(rowData, 'Matière');
        const classe = getRowDataValue(rowData, 'Classe');
        const periode = getRowDataValue(rowData, 'Période');
        const titreUnite = getRowDataValue(rowData, 'Titre de l\'unité', getRowDataValue(rowData, 'Leçon'));
        const lecon = getRowDataValue(rowData, 'Leçon');

        let dateFormatted = jour;
        const weekDates = specificWeekDateRangesNode[weekNumber];
        if (weekDates?.start && jour) {
            try {
                const weekStartDate = new Date(weekDates.start + 'T00:00:00Z');
                const lessonDateObj = getDateForDayNameNode(weekStartDate, jour);
                if (lessonDateObj) dateFormatted = formatDateFrenchNode(lessonDateObj);
            } catch (e) { console.error("[AI-PLAN] Erreur formatage date:", e); }
        }
        
        const prompt = `
Génère le contenu pour un plan de leçon basé sur les informations suivantes. Réponds de manière concise pour chaque section.
- Leçon: ${lecon || 'Non spécifié'}
- Matière: ${matiere || 'Non spécifié'}
- Classe: ${classe || 'Non spécifié'}

Structure ta réponse EXACTEMENT comme suit, en utilisant "###" comme séparateur AVANT chaque nom de section.

### METHODES:
[Décris ici les méthodes pédagogiques (ex: exposé dialogué, travail de groupe, etc.)]

### OUTILS:
[Liste ici les outils et supports didactiques (ex: manuel scolaire page X, TBI, etc.)]

### OBJECTIFS:
[Liste ici les objectifs d'apprentissage clairs (ex: - Comprendre le concept de... - Être capable d'appliquer...)]

### MINUTAGE:
[Propose un découpage temporel (ex: - Accueil (5 min) - Activité (20 min) - Synthèse (10 min))]

### CONTENU:
[Décris ici les étapes clés de la leçon de manière détaillée.]

### RESSOURCES:
[Récapitule le matériel spécifique nécessaire.]

### DEVOIRS:
[Indique clairement les devoirs à faire.]

### DIFF_LENTS:
[Propose des stratégies pour les élèves en difficulté.]

### DIFF_PERFORMANTS:
[Propose des défis pour les élèves avancés.]

### DIFF_TOUS:
[Propose des stratégies générales pour tous les élèves.]
`;

        console.log("[AI-PLAN] Appel de l'API Gemini...");
        
        const result = await geminiModel.generateContent(prompt);
        const response = await result.response;
        const aiResponseText = response.text();
        
        console.log("[AI-PLAN] Réponse Gemini reçue. Début du parsing...");

        const aiGenerated = {};
        const sectionsExpected = [
            "METHODES", "OUTILS", "OBJECTIFS", "MINUTAGE", "CONTENU", 
            "RESSOURCES", "DEVOIRS", "DIFF_LENTS", "DIFF_PERFORMANTS", "DIFF_TOUS"
        ];
        
        let currentSectionName = null;
        let currentContent = [];

        aiResponseText.split('\n').forEach(line => {
            const trimmedLine = line.trim();
            let isSectionHeader = false;
            for (const section of sectionsExpected) {
                if (trimmedLine.startsWith(`### ${section}:`)) {
                    if (currentSectionName) {
                        aiGenerated[currentSectionName] = currentContent.join('\n').trim();
                    }
                    currentSectionName = section;
                    currentContent = [trimmedLine.substring(`### ${section}:`.length).trim()];
                    isSectionHeader = true;
                    break;
                }
            }
            if (!isSectionHeader && currentSectionName) {
                currentContent.push(line);
            }
        });
        if (currentSectionName) {
            aiGenerated[currentSectionName] = currentContent.join('\n').trim();
        }

        sectionsExpected.forEach(section => {
            if (!aiGenerated[section]) {
                aiGenerated[section] = `(Non généré)`;
                console.warn(`[AI-PLAN] Section "${section}" manquante.`);
            }
        });
        
        console.log("[AI-PLAN] Parsing terminé. Préparation du document Word...");

        const AI_WORD_TEMPLATE_URL = 'https://cdn.glitch.global/d411e70d-81bc-41b6-902e-a5403e356bac/Plan_de_le%C3%A7on_modele.docx?v=1730495303423';
        const templateResponse = await fetch(AI_WORD_TEMPLATE_URL);
        if (!templateResponse.ok) throw new Error(`Échec du téléchargement du modèle Word (${templateResponse.status})`);
        const templateBuffer = Buffer.from(await templateResponse.arrayBuffer());

        const zip = new PizZip(templateBuffer);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => "" });

        const templateData = {
          enseignant: enseignant,
          date: dateFormatted,
          semaine: `Semaine ${weekNumber}`,
          matiere: matiere,
          classe: classe,
          seance: periode, 
          jour: jour,
          unite: titreUnite,
          lecon: lecon,
          methodes: aiGenerated.METHODES,
          outils: aiGenerated.OUTILS,
          objectifs: aiGenerated.OBJECTIFS,
          minutage: aiGenerated.MINUTAGE,
          contenu: aiGenerated.CONTENU,
          ressources: aiGenerated.RESSOURCES,
          devoirs: aiGenerated.DEVOIRS,
          diff_lents: aiGenerated.DIFF_LENTS,
          diff_performants: aiGenerated.DIFF_PERFORMANTS,
          diff_tous: aiGenerated.DIFF_TOUS,
        };

        doc.render(templateData);

        const buffer = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

        const safeClasseName = classe.replace(/[^a-z0-9]/gi, '_').replace(/_+/g, '_');
        const safeMatiereName = matiere.replace(/[^a-z0-9]/gi, '_').replace(/_+/g, '_');
        const filename = `Plan_Lecon_IA_S${weekNumber}_${safeClasseName}_${safeMatiereName}.docx`;

        console.log(`[AI-PLAN] Envoi du document Word: ${filename}`);
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buffer);

    } catch (error) {
        console.error('❌ Erreur serveur /generate-ai-lesson-plan:', error);
        let errorMessage = 'Erreur interne serveur (IA).';
        if (error.message) {
            errorMessage = error.message;
        }
        if (!res.headersSent) {
            res.status(500).json({
                message: errorMessage,
                errorDetails: process.env.NODE_ENV !== 'production' ? error.toString() : undefined
            });
        } else {
            console.error("[AI-PLAN] Headers déjà envoyés.");
        }
    }
    console.log('--- Fin /generate-ai-lesson-plan ---');
});


// --- Démarrage Serveur ---
server.listen(PORT, () => {
    console.log(`🚀 Serveur Express démarré sur http://localhost:${PORT}`);
    if (WORD_TEMPLATE_URL) {
        console.log(`   URL modèle Word: ${WORD_TEMPLATE_URL}`);
    } else {
        console.warn('⚠️ WORD_TEMPLATE_URL non défini dans le fichier .env. La génération de documents Word échouera.');
    }
});

// --- Gestionnaires Erreurs Globaux ---
process.on('uncaughtException', (error, origin) => { console.error(`❌ ERREUR NON CAPTURÉE! Origine: ${origin}`); console.error(error); });
process.on('unhandledRejection', (reason, promise) => { console.error('❌ REJET PROMESSE NON GÉRÉ!'); console.error('Promise:', promise, 'Raison:', reason); });
