# ExtractBankReportToExcel

Application WPF qui lit les relevés PDF du dossier `PdfFiles` et affiche les opérations bancaires dans un tableau. Les opérations reconnues peuvent être exportées vers un classeur Excel `.xlsx`.

## Utilisation

1. Copiez les relevés PDF dans `ExtractBankReportToExcel/PdfFiles`.
2. Lancez l’application. Elle crée le dossier `PdfFiles` dans son dossier d’exécution s’il n’existe pas et charge les fichiers PDF au démarrage.
3. Utilisez **Actualiser** après avoir ajouté ou remplacé des relevés.
4. Cliquez sur **Exporter en Excel (.xlsx)** et choisissez où enregistrer le classeur.

Le tableau contient les dates d’opération et de valeur, le libellé, les montants débit/crédit ainsi que le nom du PDF source. Les PDF illisibles ou dont le format ne contient pas une table d’opérations reconnue sont signalés dans un avertissement.

L’extraction actuelle reconnaît les tables comportant les colonnes date, opération, débit et crédit. Les PDF numérisés sans couche texte ne peuvent pas être lus sans OCR.
