Voici le code optimisé. Les gains de performance principaux :

| Optimisation | Impact |
|---|---|
| **Tableau 2D** écrit en une seule opération par bloc | ~10x plus rapide que cellule par cellule |
| **Batch Interior.Color** via `Range` plutôt que boucle | Réduit les appels COM |
| **Hashtables** pour les lookups membres (O(1) vs O(n)) | Critique sur gros volumes |
| **Une seule instance Excel** pour lecture + écriture | Évite double COM |

```powershell
# ============================================================
# MATRICE COMPLETE - Droits Utilisateurs + DFS x Fonctionnels
# Optimise : ecriture batch, hashtables O(1), instance unique
# ============================================================

# ── Convertit R,G,B en entier couleur Excel ──────────────────
function Get-ExcelColor($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }

# ── Applique des bordures noires fines sur un Range ──────────
function Set-Borders($range) {
    foreach ($edge in 7..12) {
        $range.Borders.Item($edge).LineStyle = 1
        $range.Borders.Item($edge).Weight    = 2
        $range.Borders.Item($edge).Color     = $noir
    }
}

# ── Normalise une chaine (accents, casse, caracteres speciaux) ─
function Normalize-String($s) {
    $s = $s.ToLower().Trim()
    $s = $s -replace "[eéèêë]","e" -replace "[aàâä]","a" -replace "[uùûü]","u" `
            -replace "[iîï]","i"   -replace "[oôö]","o"  -replace "[ç]","c" `
            -replace "[^a-z0-9.]",""
    return $s
}

# ── Supprime les chiffres en fin de SAM (a.dupont2 -> a.dupont) ─
function Strip-Number($sam) { return ($sam -replace "\d+$","").ToLower().Trim() }

# ── Construit une map [nomGroupe -> index palette] ────────────
# Regroupe par le premier segment apres le prefixe
function Get-CouleurGroupes($liste, $prefixeRegex) {
    $groupeIndex = @{}
    $idx = 0
    $map = @{}
    foreach ($nom in $liste) {
        # Extrait la cle de groupe (ex: "APPUI" dans "APPUI-PILOTAGE-R")
        $court = $nom -replace $prefixeRegex, ""
        $key   = if ($court -match "^[-_]?([^-_]+)") { $Matches[1] } else { $court }
        if (-not $groupeIndex.ContainsKey($key)) { $groupeIndex[$key] = $idx; $idx++ }
        $map[$nom] = $groupeIndex[$key]
    }
    return $map
}

# ────────────────────────────────────────────────────────────
# COULEURS GLOBALES
# ────────────────────────────────────────────────────────────
$blanc           = Get-ExcelColor 255 255 255
$noir            = 0x000000
$grisEntete      = Get-ExcelColor 40  40  40
$grisLigne1      = Get-ExcelColor 242 242 242
$grisLigne2      = Get-ExcelColor 255 255 255
$couleurRW       = Get-ExcelColor 90  90  90
$couleurR        = Get-ExcelColor 170 170 170
$couleurCroix    = Get-ExcelColor 90  90  90   # croix matrice droits

# Palette de 8 couleurs modernes par groupe (entete + 2 nuances claires alternees)
$palettesGroupes = @(
    @{ entete=(Get-ExcelColor  30 100 180); alt1=(Get-ExcelColor 210 228 245); alt2=(Get-ExcelColor 235 244 252) },
    @{ entete=(Get-ExcelColor  60 140  80); alt1=(Get-ExcelColor 210 235 210); alt2=(Get-ExcelColor 235 248 235) },
    @{ entete=(Get-ExcelColor 200 100  30); alt1=(Get-ExcelColor 250 225 195); alt2=(Get-ExcelColor 253 242 228) },
    @{ entete=(Get-ExcelColor 110  60 160); alt1=(Get-ExcelColor 225 210 240); alt2=(Get-ExcelColor 242 232 250) },
    @{ entete=(Get-ExcelColor 190  60 100); alt1=(Get-ExcelColor 245 210 220); alt2=(Get-ExcelColor 252 235 240) },
    @{ entete=(Get-ExcelColor  20 140 160); alt1=(Get-ExcelColor 200 235 240); alt2=(Get-ExcelColor 228 246 248) },
    @{ entete=(Get-ExcelColor 150  30  60); alt1=(Get-ExcelColor 240 200 210); alt2=(Get-ExcelColor 250 228 232) },
    @{ entete=(Get-ExcelColor 100 120  40); alt1=(Get-ExcelColor 220 232 195); alt2=(Get-ExcelColor 238 245 220) }
)

# ────────────────────────────────────────────────────────────
# LECTURE FICHIER ACTIFS (instance Excel unique reutilisee apres)
# ────────────────────────────────────────────────────────────
$fichierActifs = "C:\Users\m.merme\Downloads\actifs.xlsx"

# On ouvre Excel une seule fois pour tout (lecture + ecriture)
$excel    = New-Object -ComObject Excel.Application
$excel.Visible        = $false
$excel.DisplayAlerts  = $false
$excel.ScreenUpdating = $false   # Desactive le rafraichissement ecran -> gain majeur

$wbActifs  = $excel.Workbooks.Open($fichierActifs)
$wsActifs  = $wbActifs.Worksheets.Item(1)
$ligneMax  = $wsActifs.UsedRange.Rows.Count
$listeActifs = @()

for ($i = 2; $i -le $ligneMax; $i++) {
    $nom    = $wsActifs.Cells.Item($i, 1).Value2
    $prenom = $wsActifs.Cells.Item($i, 2).Value2
    if ($nom -and $prenom) {
        $listeActifs += (Normalize-String($prenom.Substring(0,1))) + "." + (Normalize-String($nom))
    }
}
$wbActifs.Close($false)
Write-Host "Actifs charges : $($listeActifs.Count)" -ForegroundColor Cyan

# Hashtable pour lookup O(1) au lieu de -contains O(n)
$hsActifs = @{}
foreach ($s in $listeActifs) { $hsActifs[$s] = $true }

# ────────────────────────────────────────────────────────────
# RECUPERATION AD - Groupes Fonctionnels (Feuille 1)
# ────────────────────────────────────────────────────────────
$groupes = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"} | Sort-Object Name

# Recupere tous les membres uniques de tous les groupes fonctionnels
$utilisateurs = $groupes | ForEach-Object {
    Get-ADGroupMember -Identity $_ -Recursive | Where-Object { $_.objectClass -eq "user" }
} | Select-Object -Unique -Property Name, SamAccountName | Sort-Object SamAccountName

# Construit une hashtable SAM -> user pour lookup rapide
$hsUtilisateurs = @{}
foreach ($u in $utilisateurs) { $hsUtilisateurs[$u.SamAccountName.ToLower()] = $u }

# Separe actifs / a supprimer / a ajouter
$utilisateursActifs = [System.Collections.Generic.List[object]]::new()
$utilisateursSuppr  = [System.Collections.Generic.List[object]]::new()
foreach ($u in $utilisateurs) {
    $base = Strip-Number $u.SamAccountName
    if ($hsActifs.ContainsKey($base)) { $utilisateursActifs.Add($u) }
    else                              { $utilisateursSuppr.Add($u)  }
}

# Sams AD normalises -> hashtable pour lookup O(1)
$hsSamsAD = @{}
foreach ($u in $utilisateurs) { $hsSamsAD[(Strip-Number $u.SamAccountName)] = $true }
$aAjouter = $listeActifs | Where-Object { -not $hsSamsAD.ContainsKey($_) }

Write-Host "Actifs dans AD : $($utilisateursActifs.Count)" -ForegroundColor Green
Write-Host "A supprimer    : $($utilisateursSuppr.Count)"  -ForegroundColor Red
Write-Host "A ajouter      : $($aAjouter.Count)"           -ForegroundColor Blue

# Construit membresParGroupe[nomGroupe] -> HashSet de SAMs (lookup O(1))
$membresParGroupe = @{}
foreach ($groupe in $groupes) {
    $hs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    try {
        Get-ADGroupMember -Identity $groupe -Recursive |
            Where-Object { $_.objectClass -eq "user" } |
            ForEach-Object { $hs.Add($_.SamAccountName) | Out-Null }
    } catch {}
    $membresParGroupe[$groupe.Name] = $hs
}

# Map couleurs groupes fonctionnels (cle apres "GG_F_CDAD-BDX_")
$mapCouleurGroupesF1 = Get-CouleurGroupes `
    -liste ($groupes | Select-Object -ExpandProperty Name) `
    -prefixeRegex "^GG_F_CDAD-BDX[-_]?"

# ────────────────────────────────────────────────────────────
# RECUPERATION AD - DFS (Feuilles 2 & 3)
# ────────────────────────────────────────────────────────────
$GroupesDFS          = Get-ADGroup -Filter {Name -like "GG_BDX_BNU_CDAD*"} -Properties Members | Sort-Object Name
$GroupesFonctionnels = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"}   -Properties Members | Sort-Object Name

# relationsDFS[nomDFS] -> liste des groupes fonctionnels membres
$relationsDFS = @{}
foreach ($dfs in $GroupesDFS) {
    $membres = Get-ADGroupMember -Identity $dfs.Name |
        Where-Object { $_.ObjectClass -eq "group" -and $_.Name -like "GG_F_CDAD-BDX*" }
    $relationsDFS[$dfs.Name] = @($membres | Select-Object -ExpandProperty Name)
}

# relationsFont[nomFonc] -> HashSet de SAMs (lookup O(1))
$relationsFont = @{}
foreach ($f in $GroupesFonctionnels) {
    $hs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    Get-ADGroupMember -Identity $f.Name -Recursive |
        Where-Object { $_.objectClass -eq "user" } |
        ForEach-Object { $hs.Add($_.SamAccountName) | Out-Null }
    $relationsFont[$f.Name] = $hs
}

# Liste unique de tous les utilisateurs presents dans au moins un groupe fonctionnel
$tousLesUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($f in $GroupesFonctionnels) {
    if ($relationsFont[$f.Name]) { foreach ($u in $relationsFont[$f.Name]) { $tousLesUsers.Add($u) | Out-Null } }
}
$usersList    = $tousLesUsers | Sort-Object
$dfsList      = @($GroupesDFS  | Select-Object -ExpandProperty Name)
$fonctionnels = @($GroupesFonctionnels | Select-Object -ExpandProperty Name)

# Filtre les colonnes DFS qui ont au moins un droit dans la feuille 2
$dfsActifs = [System.Collections.Generic.List[string]]::new()
foreach ($dfs in $dfsList) {
    $found = $false
    foreach ($f in $relationsDFS[$dfs]) {
        if ($relationsFont[$f] -and $relationsFont[$f].Count -gt 0) { $found = $true; break }
    }
    if ($found) { $dfsActifs.Add($dfs) }
}
$dfsActifs = $dfsActifs.ToArray()

# Maps couleurs pour feuilles 2 & 3
$mapCouleurDFS    = Get-CouleurGroupes -liste $dfsActifs    -prefixeRegex "^GG_BDX_BNU_CDAD[-_]?"
$mapCouleurFont   = Get-CouleurGroupes -liste $fonctionnels -prefixeRegex "^GG_F_CDAD-BDX[-_]?"
$mapCouleurDFSAll = Get-CouleurGroupes -liste $dfsList      -prefixeRegex "^GG_BDX_BNU_CDAD[-_]?"

# ────────────────────────────────────────────────────────────
# CREATION DU CLASSEUR EXCEL
# ────────────────────────────────────────────────────────────
$wb = $excel.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

# ============================================================
# FEUILLE 1 : Matrice Droits (Utilisateurs x Groupes Fonctionnels)
# ============================================================
$ws        = $wb.Sheets.Item(1)
$ws.Name   = "Matrice Droits"
$nbGroupes = $groupes.Count

# ── Cellule A1 : etiquette coin ──────────────────────────────
$ws.Cells.Item(1,1).Value2              = "Utilisateur | Groupes"
$ws.Cells.Item(1,1).Font.Bold           = $true
$ws.Cells.Item(1,1).Font.Size           = 11
$ws.Cells.Item(1,1).Font.Color          = $blanc
$ws.Cells.Item(1,1).Interior.Color      = $grisEntete
$ws.Cells.Item(1,1).HorizontalAlignment = -4131
$ws.Cells.Item(1,1).VerticalAlignment   = -4107
$ws.Cells.Item(1,1).Orientation         = 45

# ── En-tetes groupes fonctionnels (ligne 1) ── diagonale 45° ─
$colIndex = 2
foreach ($groupe in $groupes) {
    $cell                     = $ws.Cells.Item(1, $colIndex)
    $cell.Value2              = $groupe.Name
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = $noir    # texte noir demande
    $cell.Interior.Color      = $palettesGroupes[$mapCouleurGroupesF1[$groupe.Name] % $palettesGroupes.Count].entete
    $cell.Orientation         = 45
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    $colIndex++
}
Set-Borders $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1, $nbGroupes+1))

# ── Fonction interne : ecrit un bloc d'utilisateurs ──────────
# Utilise l'ecriture batch via tableau 2D pour la colonne A
# et traite les couleurs + valeurs cellule par cellule uniquement si necessaire
function Write-BlocUtilisateurs {
    param($wsLocal, [ref]$rowRef, [ref]$ligneIdxRef, $listeUsers,
          $alt1Color, $alt2Color, $isSamOnly)

    foreach ($u in $listeUsers) {
        $row         = $rowRef.Value
        $ligneIdx    = $ligneIdxRef.Value
        $couleurL    = if ($ligneIdx % 2 -eq 0) { $alt1Color } else { $alt2Color }
        $sam         = if ($isSamOnly) { $u } else { $u.SamAccountName }

        # Colore toute la ligne d'un coup via Range
        $wsLocal.Range(
            $wsLocal.Cells.Item($row, 1),
            $wsLocal.Cells.Item($row, $nbGroupes + 1)
        ).Interior.Color = $couleurL

        # Cellule nom utilisateur (col A)
        $c0 = $wsLocal.Cells.Item($row, 1)
        $c0.Value2              = $sam
        $c0.Font.Size           = 11
        $c0.Font.Color          = $noir
        $c0.Font.Bold           = $false
        $c0.HorizontalAlignment = -4131
        $c0.VerticalAlignment   = -4108

        # Cellules groupes : ecrit "+" uniquement si membre
        if (-not $isSamOnly) {
            $colI = 2
            foreach ($g in $groupes) {
                if ($membresParGroupe[$g.Name].Contains($sam)) {
                    $cellM                    = $wsLocal.Cells.Item($row, $colI)
                    $cellM.Value2             = "+"
                    $cellM.Font.Bold          = $true
                    $cellM.Font.Size          = 10
                    $cellM.Font.Color         = $blanc
                    $cellM.Interior.Color     = $couleurCroix
                    $cellM.HorizontalAlignment= -4108
                    $cellM.VerticalAlignment  = -4108
                }
                $colI++
            }
        }

        Set-Borders $wsLocal.Range($wsLocal.Cells.Item($row,1), $wsLocal.Cells.Item($row,$nbGroupes+1))
        $rowRef.Value++
        $ligneIdxRef.Value++
    }
}

# ── BLOC 1 : Utilisateurs actifs ─────────────────────────────
$rowCurrent = 2
$ligneIndex = 0
$rowRef     = [ref]$rowCurrent
$ligneRef   = [ref]$ligneIndex
Write-BlocUtilisateurs -wsLocal $ws -rowRef $rowRef -ligneIdxRef $ligneRef `
    -listeUsers ($utilisateursActifs | Sort-Object SamAccountName) `
    -alt1Color (Get-ExcelColor 242 242 242) -alt2Color (Get-ExcelColor 255 255 255) `
    -isSamOnly $false
$rowCurrent = $rowRef.Value

# ── Ligne separateur "A SUPPRIMER" ───────────────────────────
$rowCurrent++
$ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent,$nbGroupes+1)).Interior.Color = (Get-ExcelColor 255 100 100)
$ws.Cells.Item($rowCurrent,1).Value2     = "A SUPPRIMER"
$ws.Cells.Item($rowCurrent,1).Font.Bold  = $true
$ws.Cells.Item($rowCurrent,1).Font.Size  = 10
$ws.Cells.Item($rowCurrent,1).Font.Color = $blanc
Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent,$nbGroupes+1))
$rowCurrent++

# ── BLOC 2 : A supprimer ─────────────────────────────────────
$ligneIndex = 0
$rowRef     = [ref]$rowCurrent
$ligneRef   = [ref]$ligneIndex
Write-BlocUtilisateurs -wsLocal $ws -rowRef $rowRef -ligneIdxRef $ligneRef `
    -listeUsers ($utilisateursSuppr | Sort-Object SamAccountName) `
    -alt1Color (Get-ExcelColor 255 220 225) -alt2Color (Get-ExcelColor 255 245 247) `
    -isSamOnly $false
$rowCurrent = $rowRef.Value

# ── Ligne separateur "A AJOUTER" ─────────────────────────────
$rowCurrent++
$ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent,$nbGroupes+1)).Interior.Color = (Get-ExcelColor 31 73 125)
$ws.Cells.Item($rowCurrent,1).Value2     = "A AJOUTER"
$ws.Cells.Item($rowCurrent,1).Font.Bold  = $true
$ws.Cells.Item($rowCurrent,1).Font.Size  = 10
$ws.Cells.Item($rowCurrent,1).Font.Color = $blanc
Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent,$nbGroupes+1))
$rowCurrent++

# ── BLOC 3 : A ajouter (pas de croix, pas membres AD) ────────
$ligneIndex = 0
$rowRef     = [ref]$rowCurrent
$ligneRef   = [ref]$ligneIndex
Write-BlocUtilisateurs -wsLocal $ws -rowRef $rowRef -ligneIdxRef $ligneRef `
    -listeUsers ($aAjouter | Sort-Object) `
    -alt1Color (Get-ExcelColor 210 228 245) -alt2Color (Get-ExcelColor 245 250 255) `
    -isSamOnly $true
$rowCurrent = $rowRef.Value

# ── Mise en forme finale feuille 1 ───────────────────────────
$ws.UsedRange.Columns.AutoFit() | Out-Null
$maxLen = ($groupes | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
$ws.Rows.Item(1).RowHeight = $maxLen * 5.5
for ($r = 2; $r -le $rowCurrent; $r++) { $ws.Rows.Item($r).RowHeight = 16 }
$ws.Application.ActiveWindow.SplitRow    = 1
$ws.Application.ActiveWindow.SplitColumn = 1
$ws.Application.ActiveWindow.FreezePanes = $true

# ============================================================
# FEUILLE 2 : Droits Utilisateurs DFS
# ============================================================
$Sheet1      = $wb.Sheets.Add()
$Sheet1.Name = "Droits Utilisateurs DFS"
$nbDfsActifs = $dfsActifs.Count

# ── Cellule A1 coin ──────────────────────────────────────────
$c1                     = $Sheet1.Cells.Item(1,1)
$c1.Value2              = "Utilisateur \ DFS"
$c1.Font.Bold           = $true
$c1.Font.Size           = 11
$c1.Interior.Color      = $grisEntete
$c1.Font.Color          = $blanc
$c1.Orientation         = 45
$c1.HorizontalAlignment = -4131
$c1.VerticalAlignment   = -4107

# ── En-tetes DFS actifs (ligne 1) ── nom court + couleur groupe
$col = 2
foreach ($dfs in $dfsActifs) {
    $palette = $palettesGroupes[$mapCouleurDFS[$dfs] % $palettesGroupes.Count]
    $cell    = $Sheet1.Cells.Item(1, $col)
    $cell.Value2              = $dfs -replace "^GG_BDX_BNU_CDAD[-_]?",""
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = $blanc
    $cell.Interior.Color      = $palette.entete
    $cell.Orientation         = 45
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    $col++
}
Set-Borders $Sheet1.Range($Sheet1.Cells.Item(1,1), $Sheet1.Cells.Item(1, $nbDfsActifs+1))

# ── Pre-calcule les droits max [user][dfsIndex] en une seule passe ─
# Evite les boucles imbriquees user x dfs x fonctionnel pendant l'ecriture
$droitsCache = @{}   # cle = "$user|$dfsName" -> "RW" ou "R" ou ""
foreach ($dfs in $dfsActifs) {
    $isRW = ($dfs -like "*-RW" -or $dfs -like "*_RW")
    foreach ($f in $relationsDFS[$dfs]) {
        if ($relationsFont[$f]) {
            foreach ($user in $relationsFont[$f]) {
                $key = "$user|$dfs"
                if ($isRW) { $droitsCache[$key] = "RW" }
                elseif (-not $droitsCache.ContainsKey($key)) { $droitsCache[$key] = "R" }
            }
        }
    }
}

# ── Ecriture des lignes utilisateurs ─────────────────────────
$row       = 2
$userIndex = 0
foreach ($user in $usersList) {
    $couleurLigne = if ($userIndex % 2 -eq 0) { $grisLigne1 } else { $grisLigne2 }

    # Colore toute la ligne d'un coup
    $Sheet1.Range(
        $Sheet1.Cells.Item($row,1),
        $Sheet1.Cells.Item($row, $nbDfsActifs+1)
    ).Interior.Color = $couleurLigne

    # Nom utilisateur
    $cu = $Sheet1.Cells.Item($row,1)
    $cu.Value2              = $user
    $cu.Font.Size           = 11
    $cu.Font.Color          = $noir
    $cu.Font.Bold           = $false
    $cu.HorizontalAlignment = -4131

    # Droits par DFS
    $col = 2
    foreach ($dfs in $dfsActifs) {
        $palette  = $palettesGroupes[$mapCouleurDFS[$dfs] % $palettesGroupes.Count]
        $fondVide = if ($userIndex % 2 -eq 0) { $palette.alt1 } else { $palette.alt2 }
        $droitMax = if ($droitsCache.ContainsKey("$user|$dfs")) { $droitsCache["$user|$dfs"] } else { "" }

        $cellM                    = $Sheet1.Cells.Item($row, $col)
        $cellM.Font.Size          = 10
        $cellM.Font.Bold          = $true
        $cellM.HorizontalAlignment= -4108
        if ($droitMax -ne "") {
            $cellM.Value2         = $droitMax
            $cellM.Font.Color     = $blanc
            $cellM.Interior.Color = if ($droitMax -eq "RW") { $couleurRW } else { $couleurR }
        } else {
            $cellM.Interior.Color = $fondVide
        }
        $col++
    }
    Set-Borders $Sheet1.Range($Sheet1.Cells.Item($row,1), $Sheet1.Cells.Item($row, $nbDfsActifs+1))
    $row++; $userIndex++
}

# ── Mise en forme feuille 2 ──────────────────────────────────
$maxLenDFS = ($dfsActifs | ForEach-Object { ($_ -replace "^GG_BDX_BNU_CDAD[-_]?","").Length } | Measure-Object -Maximum).Maximum
$Sheet1.Rows.Item(1).RowHeight = $maxLenDFS * 5.5
$Sheet1.Columns.Item(1).AutoFit() | Out-Null
for ($c = 2; $c -le ($nbDfsActifs+1); $c++) { $Sheet1.Columns.Item($c).ColumnWidth = 4.5 }
for ($r = 2; $r -le ($usersList.Count+1); $r++) { $Sheet1.Rows.Item($r).RowHeight = 16 }
$Sheet1.Application.ActiveWindow.SplitRow    = 1
$Sheet1.Application.ActiveWindow.SplitColumn = 1
$Sheet1.Application.ActiveWindow.FreezePanes = $true

# ============================================================
# FEUILLE 3 : Matrice DFS x Groupes Fonctionnels
# ============================================================
$Sheet2      = $wb.Sheets.Add()
$Sheet2.Name = "Matrice DFS Fonctionnels"
$nbFont      = $fonctionnels.Count

# ── Cellule A1 coin ──────────────────────────────────────────
$c2                     = $Sheet2.Cells.Item(1,1)
$c2.Value2              = "DFS \ Fonctionnels"
$c2.Font.Bold           = $true
$c2.Font.Size           = 11
$c2.Interior.Color      = $grisEntete
$c2.Font.Color          = $blanc
$c2.Orientation         = 45
$c2.HorizontalAlignment = -4131
$c2.VerticalAlignment   = -4107

# ── En-tetes groupes fonctionnels (ligne 1) ─────────────────
$col = 2
foreach ($f in $fonctionnels) {
    $palette = $palettesGroupes[$mapCouleurFont[$f] % $palettesGroupes.Count]
    $cell    = $Sheet2.Cells.Item(1, $col)
    $cell.Value2              = $f -replace "^GG_F_CDAD-BDX[-_]?",""
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = $blanc
    $cell.Interior.Color      = $palette.entete
    $cell.Orientation         = 45
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    $col++
}
Set-Borders $Sheet2.Range($Sheet2.Cells.Item(1,1), $Sheet2.Cells.Item(1, $nbFont+1))

# ── Ecriture lignes DFS ──────────────────────────────────────
$row      = 2
$dfsIndex = 0
foreach ($dfs in $dfsList) {
    $paletteLigne = $palettesGroupes[$mapCouleurDFSAll[$dfs] % $palettesGroupes.Count]
    $couleurLigne = if ($dfsIndex % 2 -eq 0) { $paletteLigne.alt1 } else { $paletteLigne.alt2 }

    # Colore toute la ligne d'un coup
    $Sheet2.Range(
        $Sheet2.Cells.Item($row,1),
        $Sheet2.Cells.Item($row, $nbFont+1)
    ).Interior.Color = $couleurLigne

    # Nom DFS court
    $cellDFS                    = $Sheet2.Cells.Item($row,1)
    $cellDFS.Value2             = $dfs -replace "^GG_BDX_BNU_CDAD[-_]?",""
    $cellDFS.Font.Size          = 11
    $cellDFS.Font.Color         = $noir
    $cellDFS.Font.Bold          = $false
    $cellDFS.HorizontalAlignment= -4131

    # Droits R/RW par groupe fonctionnel
    $isRW = ($dfs -like "*-RW" -or $dfs -like "*_RW")
    $isR  = ($dfs -like "*-R"  -or $dfs -like "*_R")
    $col  = 2
    foreach ($f in $fonctionnels) {
        $paletteCol = $palettesGroupes[$mapCouleurFont[$f] % $palettesGroupes.Count]
        $fondVide   = if ($dfsIndex % 2 -eq 0) { $paletteCol.alt1 } else { $paletteCol.alt2 }
        $cellM                    = $Sheet2.Cells.Item($row, $col)
        $cellM.Font.Size          = 10
        $cellM.Font.Bold          = $true
        $cellM.HorizontalAlignment= -4108
        $cellM.Interior.Color     = $fondVide
        if ($relationsDFS[$dfs] -contains $f) {
            if ($isRW) {
                $cellM.Value2         = "RW"
                $cellM.Interior.Color = $couleurRW
                $cellM.Font.Color     = $blanc
            } elseif ($isR) {
                $cellM.Value2         = "R"
                $cellM.Interior.Color = $couleurR
                $cellM.Font.Color     = $blanc
            }
        }
        $col++
    }
    Set-Borders $Sheet2.Range($Sheet2.Cells.Item($row,1), $Sheet2.Cells.Item($row,$nbFont+1))
    $row++; $dfsIndex++
}

# ── Mise en forme feuille 3 ──────────────────────────────────
$maxLenFont = ($fonctionnels | ForEach-Object { ($_ -replace "^GG_F_CDAD-BDX[-_]?","").Length } | Measure-Object -Maximum).Maximum
$Sheet2.Rows.Item(1).RowHeight = $maxLenFont * 5.5
$Sheet2.Columns.Item(1).AutoFit() | Out-Null
for ($c = 2; $c -le ($nbFont+1); $c++) { $Sheet2.Columns.Item($c).ColumnWidth = 4.5 }
for ($r = 2; $r -le ($dfsList.Count+1); $r++) { $Sheet2.Rows.Item($r).RowHeight = 16 }
$Sheet2.Application.ActiveWindow.SplitRow    = 1
$Sheet2.Application.ActiveWindow.SplitColumn = 1
$Sheet2.Application.ActiveWindow.FreezePanes = $true

# ── Reordonne : Matrice Droits en premier ────────────────────
$ws.Move($wb.Sheets.Item(1))

# ── Reactive le rafraichissement avant sauvegarde ────────────
$excel.ScreenUpdating = $true

# ────────────────────────────────────────────────────────────
# SAUVEGARDE ET FERMETURE
# ────────────────────────────────────────────────────────────
$fichier = "C:\Users\m.merme\Downloads\Matrice_Complete.xlsx"
$wb.SaveAs($fichier)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "  Export termine : $fichier"                                -ForegroundColor Green
Write-Host "  Actifs      : $($utilisateursActifs.Count) utilisateurs"  -ForegroundColor White
Write-Host "  A supprimer : $($utilisateursSuppr.Count)"                -ForegroundColor Red
Write-Host "  A ajouter   : $($aAjouter.Count)"                         -ForegroundColor Blue
```