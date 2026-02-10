# VSFI - Very Simple First Installation

> Le script d'installation ultime pour makers, devs et créatifs sous Windows.

**Par [Artus Poulain](https://youtube.com/LesFreresPoulain)**

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

---

## Fonctionnalités

- **60+ applications** pré-configurées (dev, maker, multimédia, 3D, réseau...)
- **Interface graphique** moderne (dark theme) pour sélectionner les apps
- **3 gestionnaires de paquets** : Winget, Microsoft Store, Chocolatey
- **Fallback automatique** : si Winget échoue, Chocolatey prend le relais
- **Backup / Restore** : sauvegarde la liste de tes logiciels installés en JSON, puis restaure-la sur un autre PC
- **Import / Export** de sélections personnalisées
- **Recherche** intégrée dans l'interface
- **Détection** des applications déjà installées (skip automatique)
- **Auto-élévation** admin + mode STA automatique
- **Logs** détaillés dans `%TEMP%`

## Démarrage rapide

### Option 1 : Double-clic

1. Télécharge le repo (ou clone-le)
2. Double-clique sur **`VSFI.bat`**
3. Accepte l'élévation admin
4. Choisis ton action dans l'écran d'accueil :
   - **Installer des applications** → sélectionne et installe
   - **Sauvegarder mes logiciels (Backup)** → exporte la liste de tes apps en JSON

### Option 2 : PowerShell

```powershell
# Lancement normal (GUI)
.\VSFI.ps1

# Mode automatique (installe les apps par défaut sans GUI)
.\VSFI.ps1 -NoPrompt

# Pré-sélectionner toutes les apps
.\VSFI.ps1 -SelectAll

# Ne pas proposer le redémarrage à la fin
.\VSFI.ps1 -SkipReboot
```

## Écran d'accueil

Au lancement, un écran d'accueil te propose :

| Bouton | Action |
|--------|--------|
| **Installer des applications** | Ouvre la fenêtre de sélection des apps |
| **Sauvegarder mes logiciels** | Scanne Winget + Chocolatey et exporte un JSON |
| **Quitter** | Ferme le script |

## Backup & Restore

### Sauvegarder (sur le PC source)

1. Lance VSFI → clique **Sauvegarder mes logiciels (Backup)**
2. Le script scanne toutes les apps installées via Winget et Chocolatey
3. Choisis où enregistrer le fichier JSON

### Restaurer (sur le nouveau PC)

1. Lance VSFI → clique **Installer des applications**
2. Dans la fenêtre de sélection, clique le bouton **Importer**
3. Sélectionne le fichier JSON de backup
4. Les apps sont automatiquement cochées (les apps hors catalogue sont ajoutées dynamiquement)
5. Clique **Installer**

## Catégories d'applications

| Catégorie | Exemples |
|-----------|----------|
| Utilitaires | NanaZip, Notepad++, Everything, PowerToys, Rufus |
| Navigateur | Brave, Firefox |
| Communication | Thunderbird, Discord, Slack, Zoom, Teams |
| Dev | Git, Docker, VS Code, Windsurf, Node.js, Python |
| Terminal | Tabby, Termius, PuTTY |
| Réseau | FileZilla, RustDesk, Tailscale, Wireshark |
| Multimédia | VLC, OBS Studio, Spotify, Handbrake, FFmpeg |
| Audio | Audacity, VoiceMeeter Banana |
| Images | GIMP, Inkscape, ShareX, Caesium |
| 3D / Maker | PrusaSlicer, Bambu Studio, Fusion 360, Blender, Arduino IDE |
| Domotique | MQTT Explorer |
| Productivité | LibreOffice, iCloud, Notion, Obsidian |
| IA | LM Studio, Claude, ChatGPT |
| Gaming | Steam, Epic Games, GOG Galaxy |
| Périphériques | Logitech Options+ |

## Structure du projet

```
VSFI/
├── VSFI.bat          # Lanceur (double-clic)
├── VSFI.ps1          # Script principal PowerShell
├── README.md
├── LICENSE
└── .gitignore
```

## Prérequis

- **Windows 10/11**
- **PowerShell 5.1+** (inclus dans Windows)
- **Connexion internet**
- Les droits administrateur (auto-élévation intégrée)

## Licence

[MIT](LICENSE) - Artus Poulain
