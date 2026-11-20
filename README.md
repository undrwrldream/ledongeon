# Le Donjon — lanceur de dés et gestionnaire de combat

Page en ligne : https://undrwrldream.github.io/ledongeon/

La partie est enregistrée dans un Google Sheets (dossier `apps-script/`) : chacun joue son tour quand il veut,
et les monstres sont joués par n'importe quel appareil qui a la page ouverte.

## Installer le Google Sheets
1. Crée un Google Sheets vide (par ex. « Le Donjon — partie »).
2. Extensions → Apps Script, remplace tout par le contenu de `apps-script/Code.gs`, enregistre.
3. Déployer → Nouveau déploiement → type « Application Web » ; Exécuter en tant que : Moi ; Qui a accès : Tout le monde.
4. Autorise l'accès, puis copie l'adresse qui finit par `/exec`.

## Rôles
- **MJ** : touche le bandeau en haut → « Créer le code MJ ». Il gère le répertoire, le bestiaire et lance les combats.
- **Joueurs** : touchent le bandeau → choisissent leur héros. Ils jouent seulement pendant leur tour.
