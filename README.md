# TIA Openness Tool

Multi-version TIA Portal DataBlock Exporter — Application PowerShell + WPF utilisant l'API TIA Openness.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue?logo=powershell)
![WPF](https://img.shields.io/badge/WPF-.NET%20Framework-purple)
![License](https://img.shields.io/badge/license-MIT-green)

## Fonctionnalites

- **Multi-version** — Detecte automatiquement toutes les versions TIA Portal installees (V15 a V20+)
- **Multi-langue** — Interface en Francais, Anglais, Espagnol et Italien (drapeaux interactifs)
- **Scan automatique** — Detecte les instances TIA Portal en cours d'execution
- **Connexion directe** — Se connecte a un projet TIA Portal ouvert via l'API Openness
- **Export DataBlocks** — Genere les fichiers source `.db` (SCL) avec toutes les dependances
- **Filtrage** — Affichage avec badges colores (Global / Instance), filtre par type

## Prerequis

| Composant | Version |
|-----------|---------|
| Windows | 10 / 11 |
| PowerShell | 5.1 (Windows PowerShell) |
| TIA Portal | V15, V16, V17, V18, V19 ou V20 |
| TIA Openness | Active a l'installation de TIA Portal |

> L'utilisateur Windows doit etre membre du groupe **Siemens TIA Openness**
> (Gestion de l'ordinateur → Utilisateurs et groupes locaux)

## Installation

```bash
git clone https://github.com/JohannPx/TIA-Openness-Tool.git
```

Aucune compilation necessaire — l'application est 100% PowerShell.

## Utilisation

### Lancement

```powershell
powershell -Sta -ExecutionPolicy Bypass -File "scripts\TIA-Openness-Tool.ps1"
```

### Workflow

1. **Selectionner la version** TIA Portal dans la sidebar (auto-detectee)
2. **Scanner** les instances TIA Portal ouvertes
3. **Se connecter** a l'instance souhaitee
4. **Charger** les DataBlocks du projet
5. **Selectionner** les DBs a exporter (filtrage possible)
6. **Exporter** → fichiers `.db` generes dans un dossier horodate

### Format de sortie

Les fichiers `.db` exportes sont au format SCL (Structured Control Language) :

```
DATA_BLOCK "DB_Example"
{ S7_Optimized_Access := 'TRUE' }
VERSION : 0.1
   STRUCT
      Variable1 : Bool;
      Variable2 : Int;
   END_STRUCT;

BEGIN
   Variable1 := FALSE;
   Variable2 := 0;

END_DATA_BLOCK
```

## Architecture

```
scripts/
  TIA-Openness-Tool.ps1        # Lanceur (STA thread, console hide, module loading)
  modules/
    AppState.ps1                # Etat central ($Script:AppState)
    Localization.ps1            # Multi-langues FR/EN/ES/IT + fonction T()
    TiaVersions.ps1             # Detection multi-version DLL Siemens
    TiaConnection.ps1           # Scan, connexion, deconnexion TIA Portal
    TiaDataBlocks.ps1           # Enumeration recursive des DataBlocks
    TiaExport.ps1               # Export GenerateSource (.db)
    UIHelpers.ps1               # Composants WPF reutilisables
    UI.ps1                      # XAML, initialisation fenetre, evenements
```

### Modules

| Module | Responsabilite |
|--------|---------------|
| **AppState** | Source unique de verite (etat connexion, blocs, config) |
| **Localization** | 4 langues, fonction `T()` avec fallback FR |
| **TiaVersions** | Scan filesystem pour DLL Siemens, chargement dynamique |
| **TiaConnection** | API Openness : scan process, attach, enumeration PLC |
| **TiaDataBlocks** | Parcours recursif des groupes de blocs |
| **TiaExport** | `GenerateSource()` avec dependances, nommage `DB{N}_{Nom}.db` |
| **UIHelpers** | Badges colores, items de liste, bannieres de statut |
| **UI** | XAML WPF, drapeaux langues, navigation sidebar, evenements |

## Versions TIA Portal supportees

L'application detecte automatiquement les DLL dans :

```
C:\Program Files\Siemens\Automation\Portal V{XX}\PublicAPI\V{XX}\Siemens.Engineering.dll
```

| Version | Fichier projet | Statut |
|---------|---------------|--------|
| V15 | `.ap15` | Supporte |
| V16 | `.ap16` | Supporte |
| V17 | `.ap17` | Supporte |
| V18 | `.ap18` | Supporte |
| V19 | `.ap19` | Supporte |
| V20 | `.ap20` | Supporte |

## Auteur

**JPR** — [@JohannPx](https://github.com/JohannPx)

## Licence

MIT
