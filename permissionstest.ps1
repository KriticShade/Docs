## Analyse de ta structure AD

### Ta question : "Est-ce toujours le cas ?"

**Pas nécessairement**, voici les cas possibles :

- ✅ Le groupe `GG_F_CDAD-BDX_...` est **membre de** un groupe de permission type `...RW`, `...R`, `...M` etc.
- ⚠️ Parfois les droits sont appliqués **directement** sur le groupe fonctionnel sans groupe intermédiaire
- ⚠️ Le suffixe peut varier : `_RW`, `_R`, `_RO`, `_Modify`, `_FullControl`, `_FC` selon les conventions de ton admin

**Il faudra vérifier** la convention exacte utilisée dans ton AD avant de figer la logique.

---

## Script PowerShell

```powershell
# Prérequis : module ActiveDirectory + ImportExcel (ou export CSV)
# Install-Module ImportExcel -Scope CurrentUser

Import-Module ActiveDirectory

# ============================================================
# PARAMETRES A ADAPTER
# ============================================================
$prefixeGroupe = "GG_F_CDAD-BDX"
$outputExcel   = "C:\Temp\Matrice_Droits_AD.xlsx"

# Liste de tes utilisateurs (à adapter : OU, filtre, ou liste manuelle)
$utilisateurs = Get-ADUser -Filter * -SearchBase "OU=Utilisateurs,DC=tondomaine,DC=local" `
                -Properties MemberOf, DisplayName, SamAccountName

# ============================================================
# COLLECTE DES DONNEES
# ============================================================
$resultat = @()

foreach ($user in $utilisateurs) {

    # Récupère les groupes fonctionnels qui correspondent au préfixe
    $groupesFonctionnels = $user.MemberOf | ForEach-Object {
        Get-ADGroup $_ -Properties MemberOf, Description
    } | Where-Object { $_.Name -like "$prefixeGroupe*" }

    if (-not $groupesFonctionnels) {
        # Utilisateur sans groupe correspondant
        $resultat += [PSCustomObject]@{
            Utilisateur    = $user.DisplayName
            SamAccountName = $user.SamAccountName
            GroupeFonctionnel = "Aucun groupe correspondant"
            GroupePermission  = ""
            Permission        = ""
        }
        continue
    }

    foreach ($groupe in $groupesFonctionnels) {

        # Cherche dans "MemberOf" du groupe les groupes de permission
        $groupesPermission = $groupe.MemberOf | ForEach-Object {
            Get-ADGroup $_ -Properties Name
        }

        if (-not $groupesPermission) {
            $resultat += [PSCustomObject]@{
                Utilisateur       = $user.DisplayName
                SamAccountName    = $user.SamAccountName
                GroupeFonctionnel = $groupe.Name
                GroupePermission  = "Aucun groupe de permission trouvé"
                Permission        = "?"
            }
            continue
        }

        foreach ($gp in $groupesPermission) {

            # Extrait le suffixe de permission à la fin du nom (ex: _RW, _R, _FC)
            $permission = "Inconnue"
            if ($gp.Name -match "_(RW|R|RO|Modify|FullControl|FC|M|WRITE|READ)$") {
                $permission = $matches[1]
            }

            $resultat += [PSCustomObject]@{
                Utilisateur       = $user.DisplayName
                SamAccountName    = $user.SamAccountName
                GroupeFonctionnel = $groupe.Name
                GroupePermission  = $gp.Name
                Permission        = $permission
            }
        }
    }
}

# ============================================================
# EXPORT
# ============================================================

# Option 1 : Export Excel (nécessite le module ImportExcel)
$resultat | Export-Excel -Path $outputExcel `
            -WorksheetName "Matrice Droits" `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow `
            -TableStyle Medium6

Write-Host "✅ Export Excel terminé : $outputExcel" -ForegroundColor Green

# Option 2 : Export CSV si pas de module ImportExcel (décommenter)
# $resultat | Export-Csv -Path "C:\Temp\Matrice_Droits_AD.csv" -NoTypeInformation -Encoding UTF8
```

---

## Ce que tu obtiens dans Excel

| Utilisateur | SamAccountName | GroupeFonctionnel | GroupePermission | Permission |
|---|---|---|---|---|
| Jean Dupont | jdupont | GG_F_CDAD-BDX_Finance | GG_F_CDAD-BDX_Finance_RW | RW |
| Jean Dupont | jdupont | GG_F_CDAD-BDX_RH | GG_F_CDAD-BDX_RH_R | R |

---

## Points à vérifier avant de lancer

> 1. **Adapter le `SearchBase`** avec ton OU réelle
> 2. **Vérifier les suffixes** de permission exacts dans ton AD
> 3. **Droits AD** : nécessite un compte avec lecture sur l'AD
> 4. Le module `ImportExcel` : `Install-Module ImportExcel -Scope CurrentUser`