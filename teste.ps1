# ============================================================
# MATRICE COMPLETE - Droits Utilisateurs + DFS x Fonctionnels
# ============================================================

function Get-ExcelColor($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }

# ============================================================
# COULEURS GLOBALES
# ============================================================
$blanc       = Get-ExcelColor 255 255 255
$noir        = 0x000000
$grisEntete  = Get-ExcelColor 40 40 40
$grisLigne1  = Get-ExcelColor 242 242 242
$grisLigne2  = Get-ExcelColor 255 255 255
$couleurRW   = Get-ExcelColor 90  90  90
$couleurR    = Get-ExcelColor 170 170 170

# Palette couleurs par groupe (en-tetes)
$palettesGroupes = @(
    @{ entete = (Get-ExcelColor  30 100 180); alt1 = (Get-ExcelColor 210 228 245); alt2 = (Get-ExcelColor 235 244 252) },
    @{ entete = (Get-ExcelColor  60 140  80); alt1 = (Get-ExcelColor 210 235 210); alt2 = (Get-ExcelColor 235 248 235) },
    @{ entete = (Get-ExcelColor 200 100  30); alt1 = (Get-ExcelColor 250 225 195); alt2 = (Get-ExcelColor 253 242 228) },
    @{ entete = (Get-ExcelColor 110  60 160); alt1 = (Get-ExcelColor 225 210 240); alt2 = (Get-ExcelColor 242 232 250) },
    @{ entete = (Get-ExcelColor 190  60 100); alt1 = (Get-ExcelColor 245 210 220); alt2 = (Get-ExcelColor 252 235 240) },
    @{ entete = (Get-ExcelColor  20 140 160); alt1 = (Get-ExcelColor 200 235 240); alt2 = (Get-ExcelColor 228 246 248) },
    @{ entete = (Get-ExcelColor 150  30  60); alt1 = (Get-ExcelColor 240 200 210); alt2 = (Get-ExcelColor 250 228 232) },
    @{ entete = (Get-ExcelColor 100 120  40); alt1 = (Get-ExcelColor 220 232 195); alt2 = (Get-ExcelColor 238 245 220) }
)

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Normalize-String($s) {
    $s = $s.ToLower().Trim()
    $s = $s -replace "[eéèêë]","e" -replace "[aàâä]","a" -replace "[uùûü]","u" `
            -replace "[iîï]","i"   -replace "[oôö]","o"  -replace "[ç]","c" `
            -replace "[^a-z0-9.]",""
    return $s
}

function Strip-Number($sam) { return ($sam -replace "\d+$","").ToLower().Trim() }

function Set-Borders($range) {
    foreach ($edge in @(7,8,9,10,11,12)) {
        $range.Borders.Item($edge).LineStyle = 1
        $range.Borders.Item($edge).Weight    = 2
        $range.Borders.Item($edge).Color     = $noir
    }
}

function Write-BorderCell($cell) {
    foreach ($edge in @(7,8,9,10)) {
        $cell.Borders.Item($edge).LineStyle = 1
        $cell.Borders.Item($edge).Weight    = 2
        $cell.Borders.Item($edge).Color     = $noir
    }
}

function Write-SeparatorRow($ws, $row, $texte, $couleurFond, $couleurTexte, $nbGroupes) {
    for ($c = 1; $c -le ($nbGroupes + 1); $c++) {
        $ws.Cells.Item($row, $c).Interior.Color = $couleurFond
    }
    $ws.Cells.Item($row, 1).Value2     = $texte
    $ws.Cells.Item($row, 1).Font.Bold  = $true
    $ws.Cells.Item($row, 1).Font.Size  = 10
    $ws.Cells.Item($row, 1).Font.Color = $couleurTexte
}

function Write-UserRow($ws, $row, $label, $couleurLigne, $couleurCroix, $groupes, $membresParGroupe, $samComplet) {
    for ($c = 1; $c -le ($groupes.Count + 1); $c++) {
        $ws.Cells.Item($row, $c).Interior.Color = $couleurLigne
    }
    $cell0 = $ws.Cells.Item($row, 1)
    $cell0.Value2              = $label
    $cell0.Font.Size           = 11
    $cell0.Font.Color          = $noir
    $cell0.Font.Bold           = $false
    $cell0.HorizontalAlignment = -4131
    $cell0.VerticalAlignment   = -4108
    Write-BorderCell $cell0

    $colIdx = 2
    foreach ($groupe in $groupes) {
        $cell = $ws.Cells.Item($row, $colIdx)
        if ($samComplet -and ($membresParGroupe[$groupe.Name] -contains $samComplet.ToLower())) {
            $cell.Value2              = "+"
            $cell.Font.Bold           = $true
            $cell.Font.Size           = 10
            $cell.Font.Color          = $blanc
            $cell.Interior.Color      = $couleurCroix
            $cell.HorizontalAlignment = -4108
            $cell.VerticalAlignment   = -4108
        } else {
            $cell.Interior.Color      = $couleurLigne
            $cell.HorizontalAlignment = -4108
            $cell.VerticalAlignment   = -4108
        }
        Write-BorderCell $cell
        $colIdx++
    }
}

function Get-GroupeKey($nom, $prefixeRegex) {
    $court = $nom -replace $prefixeRegex, ""
    if ($court -match "^[-_]?([^-_]+)") { return $Matches[1] }
    return $court
}

function Get-CouleurGroupes($liste, $prefixeRegex) {
    $groupeIndex = @{}
    $idx = 0
    $map = @{}
    foreach ($nom in $liste) {
        $key = Get-GroupeKey -nom $nom -prefixeRegex $prefixeRegex
        if (-not $groupeIndex.ContainsKey($key)) { $groupeIndex[$key] = $idx; $idx++ }
        $map[$nom] = $groupeIndex[$key]
    }
    return $map
}

# ============================================================
# LECTURE FICHIER ACTIFS
# ============================================================
$fichierActifs = "C:\Users\m.merme\Downloads\actifs.xlsx"
$excelLecture  = New-Object -ComObject Excel.Application
$excelLecture.Visible = $false
$wbActifs = $excelLecture.Workbooks.Open($fichierActifs)
$wsActifs = $wbActifs.Worksheets.Item(1)
$ligneMax = $wsActifs.UsedRange.Rows.Count
$listeActifs = @()

for ($i = 2; $i -le $ligneMax; $i++) {
    $nom    = $wsActifs.Cells.Item($i, 1).Value2
    $prenom = $wsActifs.Cells.Item($i, 2).Value2
    if ($nom -and $prenom) {
        $sam = (Normalize-String($prenom.Substring(0,1))) + "." + (Normalize-String($nom))
        $listeActifs += $sam
    }
}
$wbActifs.Close($false)
$excelLecture.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelLecture) | Out-Null
Write-Host "Actifs charges : $($listeActifs.Count)" -ForegroundColor Cyan

# ============================================================
# RECUPERATION AD - Groupes Fonctionnels
# ============================================================
$groupes = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"} | Sort-Object Name

$utilisateurs = $groupes | ForEach-Object {
    Get-ADGroupMember -Identity $_ -Recursive | Where-Object { $_.objectClass -eq "user" }
} | Sort-Object SamAccountName | Select-Object -Unique -Property Name, SamAccountName

$utilisateursActifs = $utilisateurs | Where-Object {
    $listeActifs -contains (Strip-Number $_.SamAccountName)
}
$utilisateursSuppr = $utilisateurs | Where-Object {
    $listeActifs -notcontains (Strip-Number $_.SamAccountName)
}
$samsADNormalises = $utilisateurs | ForEach-Object { Strip-Number $_.SamAccountName }
$aAjouter = $listeActifs | Where-Object { $samsADNormalises -notcontains $_ }

Write-Host "Actifs dans AD : $($utilisateursActifs.Count)" -ForegroundColor Green
Write-Host "A supprimer    : $($utilisateursSuppr.Count)"  -ForegroundColor Red
Write-Host "A ajouter      : $($aAjouter.Count)"           -ForegroundColor Blue

$membresParGroupe = @{}
foreach ($groupe in $groupes) {
    try {
        $membresParGroupe[$groupe.Name] = Get-ADGroupMember -Identity $groupe -Recursive |
            Where-Object { $_.objectClass -eq "user" } |
            ForEach-Object { $_.SamAccountName.ToLower() }
    } catch { $membresParGroupe[$groupe.Name] = @() }
}

# Carte couleurs groupes fonctionnels pour feuille 1
# On extrait la cle apres "GG_F_CDAD-BDX_" ou "GG_F_CDAD-BDX-"
$mapCouleurGroupesF1 = Get-CouleurGroupes -liste ($groupes | Select-Object -ExpandProperty Name) -prefixeRegex "^GG_F_CDAD-BDX[-_]?"

# ============================================================
# RECUPERATION AD - DFS
# ============================================================
$GroupesDFS          = Get-ADGroup -Filter {Name -like "GG_BDX_BNU_CDAD*"} -Properties Members | Sort-Object Name
$GroupesFonctionnels = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"}   -Properties Members | Sort-Object Name

$relationsDFS = @{}
foreach ($dfs in $GroupesDFS) {
    $membres = Get-ADGroupMember -Identity $dfs.Name | Where-Object {$_.ObjectClass -eq "group" -and $_.Name -like "GG_F_CDAD-BDX*"}
    $relationsDFS[$dfs.Name] = $membres | Select-Object -ExpandProperty Name
}

$relationsFont = @{}
foreach ($f in $GroupesFonctionnels) {
    $membres = Get-ADGroupMember -Identity $f.Name -Recursive | Where-Object {$_.ObjectClass -eq "user"}
    $relationsFont[$f.Name] = $membres | Select-Object -ExpandProperty SamAccountName
}

$tousLesUsers = @{}
foreach ($f in $GroupesFonctionnels) {
    if ($relationsFont[$f.Name]) {
        foreach ($u in $relationsFont[$f.Name]) { $tousLesUsers[$u] = $true }
    }
}
$usersList    = $tousLesUsers.Keys | Sort-Object
$dfsList      = $GroupesDFS.Name
$fonctionnels = $GroupesFonctionnels.Name

# Colonnes DFS actives
$dfsActifs = @()
foreach ($dfs in $dfsList) {
    $aUnDroit = $false
    foreach ($user in $usersList) {
        foreach ($f in $relationsDFS[$dfs]) {
            if ($relationsFont[$f] -contains $user) { $aUnDroit = $true; break }
        }
        if ($aUnDroit) { break }
    }
    if ($aUnDroit) { $dfsActifs += $dfs }
}

$mapCouleurDFS    = Get-CouleurGroupes -liste $dfsActifs    -prefixeRegex "^GG_BDX_BNU_CDAD[-_]?"
$mapCouleurFont   = Get-CouleurGroupes -liste $fonctionnels -prefixeRegex "^GG_F_CDAD-BDX[-_]?"
$mapCouleurDFSAll = Get-CouleurGroupes -liste $dfsList      -prefixeRegex "^GG_BDX_BNU_CDAD[-_]?"

# ============================================================
# EXCEL - CREATION CLASSEUR
# ============================================================
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }

# ============================================================
# FEUILLE 1 : Matrice Droits
# ============================================================
$ws = $wb.Sheets.Item(1)
$ws.Name = "Matrice Droits"
$nbGroupes = $groupes.Count

# Cellule A1
$ws.Cells.Item(1,1).Value2              = "Utilisateur | Groupes"
$ws.Cells.Item(1,1).Font.Bold           = $true
$ws.Cells.Item(1,1).Font.Size           = 11
$ws.Cells.Item(1,1).Font.Color          = $blanc
$ws.Cells.Item(1,1).Interior.Color      = Get-ExcelColor 40 40 40
$ws.Cells.Item(1,1).HorizontalAlignment = -4131
$ws.Cells.Item(1,1).VerticalAlignment   = -4107
$ws.Cells.Item(1,1).Orientation         = 45

# En-tetes groupes fonctionnels - couleur par groupe apres BDX_, texte NOIR
$colIndex = 2
foreach ($groupe in $groupes) {
    $cell = $ws.Cells.Item(1, $colIndex)
    $cell.Value2              = $groupe.Name
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = $noir
    $paletteIdx               = $mapCouleurGroupesF1[$groupe.Name]
    $cell.Interior.Color      = $palettesGroupes[$paletteIdx % $palettesGroupes.Count].entete
    $cell.Orientation         = 45
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    Write-BorderCell $cell
    $colIndex++
}

Set-Borders $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item(1, $nbGroupes+1))

$couleurCroixUnique = Get-ExcelColor 90 90 90

# BLOC 1 - Actifs
$rowCurrent = 2
$ligneIndex = 0
foreach ($user in ($utilisateursActifs | Sort-Object SamAccountName)) {
    $couleurLigne = if ($ligneIndex % 2 -eq 0) { Get-ExcelColor 242 242 242 } else { Get-ExcelColor 255 255 255 }
    Write-UserRow -ws $ws -row $rowCurrent -label $user.SamAccountName `
        -couleurLigne $couleurLigne -couleurCroix $couleurCroixUnique `
        -groupes $groupes -membresParGroupe $membresParGroupe -samComplet $user.SamAccountName
    Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent, $nbGroupes+1))
    $rowCurrent++; $ligneIndex++
}

# BLOC 2 - A supprimer
$rowCurrent++
Write-SeparatorRow -ws $ws -row $rowCurrent -texte "A SUPPRIMER" `
    -couleurFond (Get-ExcelColor 255 100 100) -couleurTexte $blanc -nbGroupes $nbGroupes
Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent, $nbGroupes+1))
$rowCurrent++
$ligneIndex = 0
foreach ($user in ($utilisateursSuppr | Sort-Object SamAccountName)) {
    $couleurLigne = if ($ligneIndex % 2 -eq 0) { Get-ExcelColor 255 220 225 } else { Get-ExcelColor 255 245 247 }
    Write-UserRow -ws $ws -row $rowCurrent -label $user.SamAccountName `
        -couleurLigne $couleurLigne -couleurCroix $couleurCroixUnique `
        -groupes $groupes -membresParGroupe $membresParGroupe -samComplet $user.SamAccountName
    Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent, $nbGroupes+1))
    $rowCurrent++; $ligneIndex++
}

# BLOC 3 - A ajouter
$rowCurrent++
Write-SeparatorRow -ws $ws -row $rowCurrent -texte "A AJOUTER" `
    -couleurFond (Get-ExcelColor 31 73 125) -couleurTexte $blanc -nbGroupes $nbGroupes
Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent, $nbGroupes+1))
$rowCurrent++
$ligneIndex = 0
foreach ($sam in ($aAjouter | Sort-Object)) {
    $couleurLigne = if ($ligneIndex % 2 -eq 0) { Get-ExcelColor 210 228 245 } else { Get-ExcelColor 245 250 255 }
    Write-UserRow -ws $ws -row $rowCurrent -label $sam `
        -couleurLigne $couleurLigne -couleurCroix $couleurCroixUnique `
        -groupes $groupes -membresParGroupe $membresParGroupe -samComplet $null
    Set-Borders $ws.Range($ws.Cells.Item($rowCurrent,1), $ws.Cells.Item($rowCurrent, $nbGroupes+1))
    $rowCurrent++; $ligneIndex++
}

# Mise en forme feuille 1
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
$Sheet1 = $wb.Sheets.Add()
$Sheet1.Name = "Droits Utilisateurs DFS"

$c1 = $Sheet1.Cells.Item(1,1)
$c1.Value2              = "Utilisateur \ DFS"
$c1.Font.Bold           = $true
$c1.Font.Size           = 11
$c1.Interior.Color      = $grisEntete
$c1.Font.Color          = $blanc
$c1.Orientation         = 45
$c1.HorizontalAlignment = -4131
$c1.VerticalAlignment   = -4107

$col = 2
foreach ($dfs in $dfsActifs) {
    $nomCourt   = $dfs -replace "^GG_BDX_BNU_CDAD[-_]?", ""
    $paletteIdx = $mapCouleurDFS[$dfs]
    $palette    = $palettesGroupes[$paletteIdx % $palettesGroupes.Count]
    $cell = $Sheet1.Cells.Item(1, $col)
    $cell.Value2              = $nomCourt
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = $blanc
    $cell.Interior.Color      = $palette.entete
    $cell.Orientation         = 45
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    $col++
}
Set-Borders $Sheet1.Range($Sheet1.Cells.Item(1,1), $Sheet1.Cells.Item(1, $dfsActifs.Count+1))

$row = 2
$userIndex = 0
foreach ($user in $usersList) {
    $couleurLigne = if ($userIndex % 2 -eq 0) { $grisLigne1 } else { $grisLigne2 }

    for ($c = 1; $c -le ($dfsActifs.Count + 1); $c++) {
        $Sheet1.Cells.Item($row, $c).Interior.Color = $couleurLigne
    }

    $cellUser = $Sheet1.Cells.Item($row,1)
    $cellUser.Value2              = $user
    $cellUser.Font.Bold           = $false
    $cellUser.Font.Size           = 11
    $cellUser.Font.Color          = $noir
    $cellUser.HorizontalAlignment = -4131

    $col = 2
    foreach ($dfs in $dfsActifs) {
        $paletteIdx = $mapCouleurDFS[$dfs]
        $palette    = $palettesGroupes[$paletteIdx % $palettesGroupes.Count]
        $fondVide   = if ($userIndex % 2 -eq 0) { $palette.alt1 } else { $palette.alt2 }
        $droitMax   = ""
        foreach ($f in $relationsDFS[$dfs]) {
            if ($relationsFont[$f] -contains $user) {
                if ($dfs -like "*-RW" -or $dfs -like "*_RW") { $droitMax = "RW"; break }
                elseif ($droitMax -ne "RW") { $droitMax = "R" }
            }
        }
        $cellM = $Sheet1.Cells.Item($row, $col)
        $cellM.Font.Size           = 10
        $cellM.Font.Bold           = $true
        $cellM.HorizontalAlignment = -4108
        if ($droitMax -ne "") {
            $cellM.Value2         = $droitMax
            $cellM.Font.Color     = $blanc
            $cellM.Interior.Color = if ($droitMax -eq "RW") { $couleurRW } else { $couleurR }
        } else {
            $cellM.Interior.Color = $fondVide
        }
        $col++
    }
    Set-Borders $Sheet1.Range($Sheet1.Cells.Item($row,1), $Sheet1.Cells.Item($row, $dfsActifs.Count+1))
    $row++; $userIndex++
}

$maxLenDFS = ($dfsActifs | ForEach-Object { ($_ -replace "^GG_BDX_BNU_CDAD[-_]?","").Length } | Measure-Object -Maximum).Maximum
$Sheet1.Rows.Item(1).RowHeight = $maxLenDFS * 5.5
$Sheet1.Columns.Item(1).AutoFit() | Out-Null
for ($c = 2; $c -le ($dfsActifs.Count + 1); $c++) { $Sheet1.Columns.Item($c).ColumnWidth = 4.5 }
for ($r = 2; $r -le ($usersList.Count + 1); $r++) { $Sheet1.Rows.Item($r).RowHeight = 16 }
$Sheet1.Application.ActiveWindow.SplitRow    = 1
$Sheet1.Application.ActiveWindow.SplitColumn = 1
$Sheet1.Application.ActiveWindow.FreezePanes = $true

# ============================================================
# FEUILLE 3 : Matrice DFS x Fonctionnels
# ============================================================
$Sheet2 = $wb.Sheets.Add()
$Sheet2.Name = "Matrice DFS Fonctionnels"

$c2 = $Sheet2.Cells.Item(1,1)
$c2.Value2              = "DFS \ Fonctionnels"
$c2.Font.Bold           = $true
$c2.Font.Size           = 11
$c2.Interior.Color      = $grisEntete
$c2.Font.Color          = $blanc
$c2.Orientation         = 45
$c2.HorizontalAlignment = -4131
$c2.VerticalAlignment   = -4107

$col = 2
foreach ($f in $fonctionnels) {
    $nomCourt   = $f -replace "^GG_F_CDAD-BDX[-_]?", ""
    $paletteIdx = $mapCouleurFont[$f]
    $palette    = $palettesGroupes[$paletteIdx % $palettesGroupes.Count]
    $cell = $Sheet2.Cells.Item(1, $col)
    $cell.Value2              = $nomCourt
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = $blanc
    $cell.Interior.Color      = $palette.entete
    $cell.Orientation         = 45
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    $col++
}
Set-Borders $Sheet2.Range($Sheet2.Cells.Item(1,1), $Sheet2.Cells.Item(1, $fonctionnels.Count+1))

$row = 2
$dfsIndex = 0
foreach ($dfs in $dfsList) {
    $nomCourt        = $dfs -replace "^GG_BDX_BNU_CDAD[-_]?", ""
    $paletteIdxLigne = $mapCouleurDFSAll[$dfs]
    $paletteLigne    = $palettesGroupes[$paletteIdxLigne % $palettesGroupes.Count]
    $couleurLigne    = if ($dfsIndex % 2 -eq 0) { $paletteLigne.alt1 } else { $paletteLigne.alt2 }

    for ($c = 1; $c -le ($fonctionnels.Count + 1); $c++) {
        $Sheet2.Cells.Item($row, $c).Interior.Color = $couleurLigne
    }

    $cellDFS = $Sheet2.Cells.Item($row,1)
    $cellDFS.Value2              = $nomCourt
    $cellDFS.Font.Bold           = $false
    $cellDFS.Font.Size           = 11
    $cellDFS.Font.Color          = $noir
    $cellDFS.HorizontalAlignment = -4131

    $col = 2
    foreach ($f in $fonctionnels) {
        $paletteIdxCol = $mapCouleurFont[$f]
        $paletteCol    = $palettesGroupes[$paletteIdxCol % $palettesGroupes.Count]
        $fondVide      = if ($dfsIndex % 2 -eq 0) { $paletteCol.alt1 } else { $paletteCol.alt2 }
        $cellM = $Sheet2.Cells.Item($row, $col)
        $cellM.Font.Size           = 10
        $cellM.Font.Bold           = $true
        $cellM.HorizontalAlignment = -4108
        $cellM.Interior.Color      = $fondVide
        if ($relationsDFS[$dfs] -contains $f) {
            if ($dfs -like "*-RW" -or $dfs -like "*_RW") {
                $cellM.Value2         = "RW"
                $cellM.Interior.Color = $couleurRW
                $cellM.Font.Color     = $blanc
            } elseif ($dfs -like "*-R" -or $dfs -like "*_R") {
                $cellM.Value2         = "R"
                $cellM.Interior.Color = $couleurR
                $cellM.Font.Color     = $blanc
            }
        }
        $col++
    }
    Set-Borders $Sheet2.Range($Sheet2.Cells.Item($row,1), $Sheet2.Cells.Item($row, $fonctionnels.Count+1))
    $row++; $dfsIndex++
}

$maxLenFont = ($fonctionnels | ForEach-Object { ($_ -replace "^GG_F_CDAD-BDX[-_]?","").Length } | Measure-Object -Maximum).Maximum
$Sheet2.Rows.Item(1).RowHeight = $maxLenFont * 5.5
$Sheet2.Columns.Item(1).AutoFit() | Out-Null
for ($c = 2; $c -le ($fonctionnels.Count + 1); $c++) { $Sheet2.Columns.Item($c).ColumnWidth = 4.5 }
for ($r = 2; $r -le ($dfsList.Count + 1); $r++) { $Sheet2.Rows.Item($r).RowHeight = 16 }
$Sheet2.Application.ActiveWindow.SplitRow    = 1
$Sheet2.Application.ActiveWindow.SplitColumn = 1
$Sheet2.Application.ActiveWindow.FreezePanes = $true

# Mettre Matrice Droits en premier
$ws.Move($wb.Sheets.Item(1))

# ============================================================
# SAUVEGARDE
# ============================================================
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
