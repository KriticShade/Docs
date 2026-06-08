function Get-ExcelColor($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

$palette = @(
    (Get-ExcelColor 183 222 232),
    (Get-ExcelColor 216 228 188),
    (Get-ExcelColor 252 213 180),
    (Get-ExcelColor 230 184 183),
    (Get-ExcelColor 204 192 218)
)

# ============================================================
# LECTURE DU FICHIER EXCEL DES ACTIFS
# ============================================================
$fichierActifs = "C:\Users\m.merme\Downloads\actifs.xlsx"

$excelLecture = New-Object -ComObject Excel.Application
$excelLecture.Visible = $false
$wbActifs = $excelLecture.Workbooks.Open($fichierActifs)
$wsActifs = $wbActifs.Worksheets.Item(1)

$ligneMax = $wsActifs.UsedRange.Rows.Count
$listeActifs = @()

function Normalize-String($s) {
    $s = $s.ToLower().Trim()
    $s = $s -replace '[éèêë]','e' -replace '[àâä]','a' -replace '[ùûü]','u' `
            -replace '[îï]','i'   -replace '[ôö]','o'  -replace '[ç]','c' `
            -replace '[^a-z0-9.]',''
    return $s
}

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
Write-Host "Actifs chargés : $($listeActifs.Count)" -ForegroundColor Cyan

# ============================================================
# RÉCUPÉRATION AD
# ============================================================
$groupes = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"} | Sort-Object Name

$utilisateurs = $groupes | ForEach-Object {
    Get-ADGroupMember -Identity $_ -Recursive | Where-Object { $_.objectClass -eq "user" }
} | Sort-Object SamAccountName | Select-Object -Unique -Property Name, SamAccountName

function Strip-Number($sam) {
    return ($sam -replace '\d+$','').ToLower().Trim()
}

# ============================================================
# CROISEMENT
# ============================================================
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

# ============================================================
# PRÉ-CALCUL MEMBRES PAR GROUPE
# ============================================================
$membresParGroupe = @{}
foreach ($groupe in $groupes) {
    try {
        $membresParGroupe[$groupe.Name] = Get-ADGroupMember -Identity $groupe -Recursive |
            Where-Object { $_.objectClass -eq "user" } |
            ForEach-Object { $_.SamAccountName.ToLower() }
    } catch {
        $membresParGroupe[$groupe.Name] = @()
    }
}

# ============================================================
# EXPORT EXCEL
# ============================================================
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Matrice Droits"

$nbGroupes = $groupes.Count

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================
function Write-BorderCell($cell) {
    foreach ($edge in @(7,8,9,10)) {
        $cell.Borders.Item($edge).LineStyle = 1
        $cell.Borders.Item($edge).Weight    = 2
        $cell.Borders.Item($edge).Color     = 0x000000
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

function Write-UserRow($ws, $row, $label, $couleurCase, $couleurCroix, $groupes, $membresParGroupe, $samComplet) {
    # Colonne nom toujours blanche
    $cell0 = $ws.Cells.Item($row, 1)
    $cell0.Value2              = $label
    $cell0.Interior.Color      = 0xFFFFFF
    $cell0.Font.Size           = 10
    $cell0.Font.Color          = 0x000000
    $cell0.Font.Bold           = $false
    $cell0.HorizontalAlignment = -4131  # gauche
    $cell0.VerticalAlignment   = -4108  # centré
    Write-BorderCell $cell0

    $colIdx = 2
    foreach ($groupe in $groupes) {
        $cell = $ws.Cells.Item($row, $colIdx)
        if ($samComplet -and ($membresParGroupe[$groupe.Name] -contains $samComplet.ToLower())) {
            $cell.Value2              = "+"
            $cell.Font.Bold           = $true
            $cell.Font.Size           = 10
            $cell.Font.Color          = 0x000000
            $cell.Interior.Color      = $couleurCroix
            $cell.HorizontalAlignment = -4108  # centré
            $cell.VerticalAlignment   = -4108  # centré
        } else {
            $cell.Interior.Color      = $couleurCase
            $cell.HorizontalAlignment = -4108
            $cell.VerticalAlignment   = -4108
        }
        Write-BorderCell $cell
        $colIdx++
    }
}

# ============================================================
# EN-TÊTE LIGNE 1
# ============================================================
$ws.Cells.Item(1,1).Value2              = "Utilisateur"
$ws.Cells.Item(1,1).Font.Bold           = $true
$ws.Cells.Item(1,1).Font.Size           = 10
$ws.Cells.Item(1,1).Font.Color          = 0xFFFFFF
$ws.Cells.Item(1,1).Interior.Color      = 0x000000
$ws.Cells.Item(1,1).HorizontalAlignment = -4108
$ws.Cells.Item(1,1).VerticalAlignment   = -4108

$colIndex  = 2
$colorIndex = 0
foreach ($groupe in $groupes) {
    $cell = $ws.Cells.Item(1, $colIndex)
    $cell.Value2              = $groupe.Name
    $cell.Font.Bold           = $true
    $cell.Font.Size           = 10
    $cell.Font.Color          = 0x000000
    $cell.Interior.Color      = $palette[$colorIndex % $palette.Count]
    $cell.Orientation         = 90
    $cell.HorizontalAlignment = -4108
    $cell.VerticalAlignment   = -4107
    Write-BorderCell $cell
    $colIndex++
    $colorIndex++
}

# ============================================================
# BLOC 1 — UTILISATEURS ACTIFS
# ============================================================
$rowCurrent = 2
$ligneIndex = 0

foreach ($user in ($utilisateursActifs | Sort-Object SamAccountName)) {
    if ($ligneIndex % 2 -eq 0) {
        $couleurCase  = Get-ExcelColor 242 242 242
        $couleurCroix = Get-ExcelColor 99  99  99
    } else {
        $couleurCase  = Get-ExcelColor 225 225 225
        $couleurCroix = Get-ExcelColor 80  80  80
    }
    Write-UserRow -ws $ws -row $rowCurrent `
        -label $user.SamAccountName `
        -couleurCase  $couleurCase `
        -couleurCroix $couleurCroix `
        -groupes $groupes `
        -membresParGroupe $membresParGroupe `
        -samComplet $user.SamAccountName
    $rowCurrent++
    $ligneIndex++
}

# ============================================================
# BLOC 2 — A SUPPRIMER
# ============================================================
$rowCurrent++
Write-SeparatorRow -ws $ws -row $rowCurrent `
    -texte "⚠ A SUPPRIMER" `
    -couleurFond  (Get-ExcelColor 255 100 100) `
    -couleurTexte 0xFFFFFF `
    -nbGroupes $nbGroupes
$rowCurrent++

$ligneIndex = 0
foreach ($user in ($utilisateursSuppr | Sort-Object SamAccountName)) {
    if ($ligneIndex % 2 -eq 0) {
        $couleurCase  = Get-ExcelColor 255 220 225
        $couleurCroix = Get-ExcelColor 192 0   0
    } else {
        $couleurCase  = Get-ExcelColor 255 245 247
        $couleurCroix = Get-ExcelColor 192 0   0
    }
    Write-UserRow -ws $ws -row $rowCurrent `
        -label $user.SamAccountName `
        -couleurCase  $couleurCase `
        -couleurCroix $couleurCroix `
        -groupes $groupes `
        -membresParGroupe $membresParGroupe `
        -samComplet $user.SamAccountName
    $rowCurrent++
    $ligneIndex++
}

# ============================================================
# BLOC 3 — A AJOUTER
# ============================================================
$rowCurrent++
Write-SeparatorRow -ws $ws -row $rowCurrent `
    -texte "⚠ A AJOUTER" `
    -couleurFond  (Get-ExcelColor 31 73 125) `
    -couleurTexte 0xFFFFFF `
    -nbGroupes $nbGroupes
$rowCurrent++

$ligneIndex = 0
foreach ($sam in ($aAjouter | Sort-Object)) {
    if ($ligneIndex % 2 -eq 0) {
        $couleurCase = Get-ExcelColor 210 228 245
    } else {
        $couleurCase = Get-ExcelColor 245 250 255
    }
    Write-UserRow -ws $ws -row $rowCurrent `
        -label $sam `
        -couleurCase  $couleurCase `
        -couleurCroix 0xFFFFFF `
        -groupes $groupes `
        -membresParGroupe $membresParGroupe `
        -samComplet $null
    $rowCurrent++
    $ligneIndex++
}

# AUTOFIT GLOBAL
$ws.UsedRange.Rows.AutoFit()    | Out-Null
$ws.UsedRange.Columns.AutoFit() | Out-Null
# La ligne 1 a les noms en rotation 90° donc on recalcule - sa hauteur selon le nom le plus long (sans contrainte)
$maxLen = ($groupes | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
$ws.Rows.Item(1).RowHeight = $maxLen * 5.5
# FIGER LES VOLETS
$ws.Application.ActiveWindow.SplitRow    = 1
$ws.Application.ActiveWindow.SplitColumn = 1
$ws.Application.ActiveWindow.FreezePanes = $true

# ============================================================
# SAUVEGARDE
# ============================================================
$fichier = "C:\Users\*\Downloads\matrice_droits.xlsx"
$wb.SaveAs($fichier)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host ""
Write-Host "  Export terminé : $fichier"                               -ForegroundColor Green
Write-Host "  Actifs      : $($utilisateursActifs.Count) utilisateurs" -ForegroundColor White
Write-Host "  A supprimer : $($utilisateursSuppr.Count)"               -ForegroundColor Red
Write-Host "  A ajouter   : $($aAjouter.Count)"                        -ForegroundColor Blue
