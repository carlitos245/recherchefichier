// Écouteur d'événement pour le bouton "Rechercher"
document.getElementById('searchButton').addEventListener('click', async () => {
  await searchInFiles(); // Appelle la fonction principale
});

// Fonction principale pour lire les fichiers et rechercher le mot
async function searchInFiles() {
  const searchTerm = document.getElementById('searchTerm').value;
  const directoryInput = document.getElementById('directoryInput');
  const files = directoryInput.files;
  const resultsElement = document.getElementById('results');

  // Vérification des champs
  if (!searchTerm) {
    alert('Veuillez entrer un mot à rechercher.');
    return;
  }

  if (files.length === 0) {
    alert('Veuillez sélectionner un dossier.');
    return;
  }

  let results = `Fichiers contenant le mot '${searchTerm}':<br>`;
  let found = false;

  // Parcourir les fichiers et rechercher le mot
  for (const file of files) {
    // Ignorer les fichiers trop volumineux
    if (file.size > 10485760) {
      console.warn(`Fichier ignoré (trop volumineux) : ${file.name}`);
      continue;
    }

    // Traitement des fichiers TXT
    if (file.type === 'text/plain') {
      try {
        const content = await readFileAsText(file);
        if (content.includes(searchTerm)) {
          const fileURL = URL.createObjectURL(file);
          results += `<br><a href="${fileURL}" target="_blank">Fichier TXT : ${file.name}</a><br>`;
          found = true;
        }
      } catch (error) {
        console.error(`Erreur de lecture du fichier texte ${file.name}:`, error);
      }
    }
    // Traitement des fichiers PDF
    else if (file.type === 'application/pdf') {
      try {
        const content = await readPdfAsText(file);
        if (content.includes(searchTerm)) {
          const fileURL = URL.createObjectURL(file);
          results += `<br><a href="${fileURL}" target="_blank">Fichier PDF : ${file.name}</a><br>`;
          found = true;
        }
      } catch (error) {
        console.error(`Erreur de lecture du fichier PDF ${file.name}:`, error);
      }
    }
    // Traitement des fichiers HTML
    else if (file.type === 'text/html') {
      try {
        const content = await readFileAsText(file); // Utiliser la fonction de lecture des fichiers texte
        if (content.includes(searchTerm)) {
          const fileURL = URL.createObjectURL(file);
          results += `<br><a href="${fileURL}" target="_blank">Fichier HTML : ${file.name}</a><br>`;
          found = true;
        }
      } catch (error) {
        console.error(`Erreur de lecture du fichier HTML ${file.name}:`, error);
      }
    }
    // Traitement des fichiers Word (.docx)
    else if (file.name.endsWith('.docx')) {
      try {
        const content = await readDocxAsText(file);
        if (content.includes(searchTerm)) {
          const fileURL = URL.createObjectURL(file);
          results += `<br><a href="${fileURL}" target="_blank">Fichier Word : ${file.name}</a><br>`;
          found = true;
        }
      } catch (error) {
        console.error(`Erreur de lecture du fichier Word ${file.name}:`, error);
      }
    }
    // Traitement des fichiers Excel (.xlsx)
    else if (file.name.endsWith('.xlsx')) {
      try {
        const content = await readExcelAsText(file);
        if (content.includes(searchTerm)) {
          const fileURL = URL.createObjectURL(file);
          results += `<br><a href="${fileURL}" target="_blank">Fichier Excel : ${file.name}</a><br>`;
          found = true;
        }
      } catch (error) {
        console.error(`Erreur de lecture du fichier Excel ${file.name}:`, error);
      }
    }
  }

  // Afficher les résultats
  if (!found) {
    results += "Aucun fichier ne contient ce mot.";
  }

  resultsElement.innerHTML = results; // Utilisation de innerHTML pour ajouter des liens cliquables
}

// Fonction pour lire un fichier texte (incluant les fichiers HTML)
function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

// Fonction pour lire un fichier PDF
async function readPdfAsText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = '';
  // Parcourir les pages du PDF et extraire le texte
  for (let i = 0; i < pdf.numPages; i++) {
    const page = await pdf.getPage(i + 1);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(' ') + '\n';
  }
  return text;
}

// Fonction pour lire un fichier Word
async function readDocxAsText(file) {
  const arrayBuffer = await file.arrayBuffer();
  const mammoth = await import('mammoth'); // Importer Mammoth.js de manière dynamique
  const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
  return result.value; // Renvoie le texte brut extrait
}

// Fonction pour lire un fichier Excel
async function readExcelAsText(file) {
  const arrayBuffer = await file.arrayBuffer(); // Charger le fichier en mémoire
  const workbook = XLSX.read(arrayBuffer, { type: 'array' }); // Analyser le fichier Excel
  let text = '';

  // Parcourt toutes les feuilles du fichier
  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    text += XLSX.utils.sheet_to_csv(sheet) + '\n'; // Convertit les données en texte CSV
  });

  return text;
}

