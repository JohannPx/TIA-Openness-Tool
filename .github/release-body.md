# TIA Openness Tool {{VERSION}}

**Date de release** : {{DATE}}
**Commit** : `{{COMMIT_SHA}}`

---

## A propos de cette version

Application PowerShell + WPF pour exporter les DataBlocks d'un projet TIA Portal via l'API Openness.
Interface multilingue (FR/EN/ES/IT), detection automatique des versions TIA Portal V15 a V20.

### Dernier changement
```
{{COMMIT_MSG}}
```

---

## Telecharger

> **Fichier unique auto-contenu** : tous les modules sont integres dans le script lors du build.
> Aucune dependance externe, PowerShell 5.1 natif Windows suffit.

### Ou trouver le fichier ?

Le fichier **`TIA-Openness-Tool_latest.ps1`** se trouve dans la section **Assets** tout en bas de cette page (cliquez sur **Assets** pour deplier si necessaire).

### Lancement

1. **Telecharger** le fichier `TIA-Openness-Tool_latest.ps1` depuis les **Assets** ci-dessous
2. **Ouvrir PowerShell** : clic-droit sur le menu Demarrer → **Terminal** (ou **Windows PowerShell**)
3. **Lancer** le script :

```powershell
powershell -Sta -ExecutionPolicy Bypass -File "$HOME\Downloads\TIA-Openness-Tool_latest.ps1"
```

> `$HOME\Downloads` correspond au dossier Telechargements. Si le fichier est ailleurs, adaptez le chemin.

### Avertissement de securite Windows

Au premier lancement, Windows peut afficher un avertissement car le script provient d'Internet.
Tapez `O` puis Entree pour executer. Pour ne plus voir cet avertissement : clic-droit sur le fichier → Proprietes → cochez Debloquer → OK.

---

## Fonctionnalites

### Multi-version TIA Portal
- Detection automatique des versions V15 a V20 installees
- Scan des instances TIA Portal en cours d'execution (avec nom du projet)
- Connexion directe via l'API Openness

### Table d'echange CSV
- Format UTF-8 BOM, separateur `;`
- Colonnes : Tag, DB, Offset, Type, Description, Unite, Repere, Coef
- En-tete PLC (nom, adresse IP, TSAP)

### Export var_lst (Ewon Flexy)
- Format Latin-1, 62 colonnes compatible import Ewon
- Adresses S7 generees automatiquement (ex: `DB1F2,ISOTCP,192.168.1.100,03.02`)
- Configuration Ewon : Repere, Topic (A/B/C), Page (1-11)
- Resolution automatique des unites UNECE

### Export PcVue Architect (BETA)
- 1 fichier CSV par DB dans un dossier horodate
- Colonnes : Nom, Type, Description, Decalage, WBIT, Trame
- Word-swap automatique des offsets Bool (big-endian Siemens vers PcVue)
- Compatible import PcVue Architect

### Traitement avance
- Expansion automatique des types Array en elements individuels
- Resolution des commentaires depuis les definitions FB/UDT source
- Calcul des offsets S7 pour DBs non-optimises

### Interface
- 4 langues (FR/EN/ES/IT) avec drapeaux interactifs
- Filtrage par type de DB (Global / Instance)
- Selection individuelle ou groupee des DataBlocks

---

## Prerequis

| Composant | Minimum |
|-----------|---------|
| **Windows** | 10 / 11 |
| **PowerShell** | 5.1 (inclus dans Windows) |
| **TIA Portal** | V15, V16, V17, V18, V19 ou V20 |
| **TIA Openness** | Active a l'installation |

> L'utilisateur Windows doit etre membre du groupe **Siemens TIA Openness**
> (Gestion de l'ordinateur → Utilisateurs et groupes locaux)

---

## Support

En cas de probleme :
1. Verifiez que vous utilisez la derniere version
2. Consultez la [documentation](../../README.md)
3. Ouvrez une [issue](../../issues) avec une capture d'ecran de l'erreur

---

*Release automatique generee par GitHub Actions*
