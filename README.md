# TIA Openness Tool

Multi-version TIA Portal DataBlock Exporter — Application PowerShell + WPF utilisant l'API TIA Openness.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue?logo=powershell)
![WPF](https://img.shields.io/badge/WPF-.NET%20Framework-purple)
![License](https://img.shields.io/badge/license-MIT-green)

## Fonctionnalites

- **Multi-version** — Detecte automatiquement toutes les versions TIA Portal installees (V15 a V20+)
- **Multi-langue** — Interface en Francais, Anglais, Espagnol et Italien (drapeaux interactifs)
- **Scan automatique** — Detecte les instances TIA Portal en cours d'execution (avec nom du projet)
- **Connexion directe** — Se connecte a un projet TIA Portal ouvert via l'API Openness
- **Export CSV (Siemens)** — Table d'echange avec colonnes Tag, DB, Offset, Type, Description, Unite, Repere, Coef
- **Export var_lst (Ewon)** — Fichier 62 colonnes compatible routeur Ewon Flexy (adresses S7, types, unites UNECE)
- **Expansion des Array** — Les types `Array[lo..hi] of Type` sont eclates en elements individuels avec offsets corrects
- **Commentaires UDT** — Resolution automatique des commentaires depuis les definitions FB/UDT source
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
2. **Scanner** les instances TIA Portal ouvertes (PID + nom du projet affiche)
3. **Se connecter** a l'instance souhaitee
4. **Charger** les DataBlocks du projet
5. **Selectionner** les DBs a exporter (filtrage possible)
6. **Choisir le format** d'export : CSV (Siemens) ou var_lst (Ewon)
7. **Exporter** → fichier genere dans le dossier choisi

### Formats de sortie

#### CSV (Siemens — Table d'Echange)

Export UTF-8 avec BOM, separateur `;`. En-tete PLC (nom, IP, TSAP) suivi des colonnes :

```
Tag;DB;Offset;Type;Description;Unite;Repere;Coef
Env.Moteur1.Mode;1;0.0;Int;Mode (0=Arret/1=Auto/2=Manu);;;
Env.Moteur1.Etat;1;2.0;Int;Etat moteur;;;
```

#### var_lst (Ewon Flexy)

Export Latin-1 (ISO-8859-1), separateur `;`, 62 colonnes entre guillemets. Compatible import direct Ewon.

Configuration Ewon dans l'interface :
- **Repere** — Prefixe variable (ex: `i30`)
- **Topic** — Canal de communication : A, B ou C
- **Page** — Page Ewon : 1 a 11

Adresses S7 generees automatiquement : `DB1F2,ISOTCP,192.168.1.100,03.02`

### Expansion des Array

Les types `Array[lo..hi] of BaseType` sont automatiquement eclates :

```
Consigne_Temp[1];1;120.0;Real;;;
Consigne_Temp[2];1;124.0;Real;;;
Consigne_Temp[3];1;128.0;Real;;;
...
```

Chaque element recoit son propre offset calcule selon la taille du type de base.

### Resolution des commentaires

Pour les DB d'instance, les commentaires ne sont pas dans le XML du DB mais dans les definitions FB/UDT source.
L'outil exporte automatiquement les types references (FB + UDT) et injecte leurs commentaires dans l'export.

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
    TiaExportTable.ps1          # Export CSV (Siemens) + var_lst (Ewon)
    UIHelpers.ps1               # Composants WPF reutilisables
    UI.ps1                      # XAML, initialisation fenetre, evenements
  data/
    units.json                  # Table de correspondance unites → codes UNECE
```

### Modules

| Module | Responsabilite |
|--------|---------------|
| **AppState** | Source unique de verite (etat connexion, blocs, config export, config Ewon) |
| **Localization** | 4 langues, fonction `T()` avec fallback FR |
| **TiaVersions** | Scan filesystem pour DLL Siemens, chargement dynamique |
| **TiaConnection** | API Openness : scan process (avec nom projet), attach, enumeration PLC |
| **TiaDataBlocks** | Parcours recursif des groupes de blocs + recherche par nom/numero |
| **TiaExport** | `GenerateSource()` avec dependances, nommage `DB{N}_{Nom}.db` |
| **TiaExportTable** | Export CSV table + var_lst Ewon (parsing XML, offsets S7, expansion Array, commentaires UDT) |
| **UIHelpers** | Badges colores, items de liste, bannieres de statut |
| **UI** | XAML WPF, drapeaux langues, selecteur format, config Ewon, evenements |

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
