# 📦 TIA Openness Tool {{VERSION}}

**Date de release** : {{DATE}}
**Commit** : `{{COMMIT_SHA}}`

---

## 🎯 À propos de cette version

Application **PowerShell + WPF** pour exporter les DataBlocks d'un projet TIA Portal via l'API Openness.
Interface multilingue (FR/EN/ES/IT), détection automatique des versions TIA Portal V15 à V20, export CSV (Table d'échange Siemens), var_lst (Ewon Flexy) et PcVue Architect.

### 📝 Changements de cette version

{{CHANGELOG}}

---

## 📥 Téléchargement et installation

### 🔽 Option recommandée : Exécutable (.exe)

1. **Télécharger** le fichier **`TiaOpennessTool.exe`** depuis les **Assets** ci-dessous
2. **Double-cliquer** pour lancer

Au premier lancement :
- L'application s'installe automatiquement dans votre profil utilisateur (aucun droit administrateur requis)
- Un raccourci est créé sur le **Bureau** et dans le **Menu Démarrer**
- Les lancements suivants se font via le raccourci

**Mises à jour automatiques** : à chaque démarrage, l'application vérifie si une nouvelle version est disponible sur GitHub et se met à jour silencieusement.

> **Avertissements de sécurité au premier téléchargement/lancement :**
>
> 1. **Navigateur** (Chrome/Edge) : *"TiaOpennessTool.exe n'est pas fréquemment téléchargé"*
>    - Chrome : cliquez sur **`^`** (flèche) → **Conserver**
>    - Edge : cliquez sur **`...`** → **Conserver** → **Conserver quand même**
>
> 2. **Windows SmartScreen** : *"Windows a protégé votre ordinateur"*
>    - Cliquez sur **Plus d'infos** → **Exécuter quand même**
>
> Ces avertissements sont normaux pour un exécutable non signé et n'apparaissent qu'au premier téléchargement.

### 🔽 Option avancée : Script PowerShell (.ps1)

Pour les utilisateurs avancés ou les environnements qui bloquent les exécutables non signés :

1. **Télécharger** le fichier `TIA-Openness-Tool_latest.ps1` depuis les **Assets** ci-dessous
2. **Ouvrir PowerShell** : clic-droit sur le menu Démarrer → **Terminal** (ou **Windows PowerShell**)
3. **Lancer** :

```powershell
powershell -Sta -ExecutionPolicy Bypass -File "$HOME\Downloads\TIA-Openness-Tool_latest.ps1"
```

> 💡 Adaptez le chemin si vous avez déplacé le fichier. Le `.ps1` est auto-contenu (aucun dossier `modules/` requis à côté).

---

## ✨ Fonctionnalités principales

### 🌍 Multilingue (FR/EN/ES/IT)
- ✅ Sélection de la langue via drapeaux
- ✅ Changement instantané de toute l'interface

### 🔌 Connexion TIA Portal
- ✅ Détection automatique des instances TIA Portal ouvertes (V15 → V20)
- ✅ Chargement automatique de la bonne DLL `Siemens.Engineering`

### 📤 Export DataBlocks
- ✅ **Table d'échange CSV** (format Siemens)
- ✅ **var_lst** (Ewon Flexy)
- ✅ **PcVue Architect**
- ✅ Expansion des Array, résolution + chaînage des commentaires de structures imbriquées

---

## 📋 Configuration requise

| Composant | Minimum |
|-----------|---------|
| **Windows** | 10/11 ou Server 2016+ |
| **PowerShell** | 5.1 (inclus) |
| **TIA Portal** | V15 à V20 avec Openness installé |

> ⚠️ L'utilisateur Windows doit appartenir au groupe local **`Siemens TIA Openness`** pour autoriser l'accès Openness.

---

## 🐛 Support

En cas de problème :
1. Vérifiez que vous utilisez la dernière version
2. Consultez la [documentation](https://github.com/JohannPx/TIA-Openness-Tool#readme)
3. Ouvrez une [issue](https://github.com/JohannPx/TIA-Openness-Tool/issues) avec une capture d'écran de l'erreur

---

## ⚠️ Note importante

Cet outil est destiné à un **usage professionnel** par les équipes Clauger et leurs clients autorisés.

---

*Release automatique générée par GitHub Actions*
