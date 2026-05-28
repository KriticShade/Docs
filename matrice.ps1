function Get-ExcelColor($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

$palette = @(
    (Get-ExcelColor 183  222  232),  # Bleu
    (Get-ExcelColor 216  228  188),  # Vert
    (Get-ExcelColor 252 213   180),  # Orange
    (Get-ExcelColor 230  184  183),  # Rose foncé
    (Get-ExcelColor 204  192  218)   # Violet foncé
)

# Récupération des groupes
$groupes = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"} | Sort-Object Name

# Récupération des utilisateurs
$utilisateurs = $groupes | ForEach-Object {
    Get-ADGroupMember -Identity $_ -Recursive | Where-Object { $_.objectClass -eq "user" }
} | Sort-Object SamAccountName | Select-Object -Unique -Property Name, SamAccountName

# Construction de la matrice
$matrice = foreach ($user in $utilisateurs) {
    $userAD = Get-ADUser $user.SamAccountName -Properties MemberOf, DisplayName
    $groupesUser = $userAD.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
    $ligne = [ordered]@{ "Utilisateur" = $userAD.SamAccountName }
    foreach ($groupe in $groupes) {
        $ligne[$groupe.Name] = if ($groupesUser -contains $groupe.Name) { "+" } else { "" }
    }
    [PSCustomObject]$ligne
}

# ---- EXCEL VIA COM ----
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Matrice Droits"

$colonnes = $matrice[0].PSObject.Properties.Name

# ---- EN-TÊTE ----
for ($col = 0; $col -lt $colonnes.Count; $col++) {
    $cell = $ws.Cells.Item(1, $col + 1)
    $cell.Value2 = $colonnes[$col]
    $cell.Font.Bold = $true
    $cell.Font.Size = 7

    if ($col -eq 0) {
        $cell.Interior.Color = Get-ExcelColor 255 255 255
        $cell.Font.Color = 0xFFFFFF
        $cell.VerticalAlignment = -4107
    } else {
        $cell.Interior.Color = $palette[($col - 1) % 5]
        $cell.Font.Color = 000000      # Texte blanc sur fond foncé
        $cell.Orientation = 90
        $cell.VerticalAlignment = -4107
        $cell.HorizontalAlignment = -4108
    }
}

# Hauteur ligne en-tête réduite au minimum viable avec rotation
$ws.Rows.Item(1).RowHeight = 100

# ---- DONNÉES ----
for ($lig = 0; $lig -lt $matrice.Count; $lig++) {
    $ligne = $matrice[$lig]
    $excelLig = $lig + 2

    for ($col = 0; $col -lt $colonnes.Count; $col++) {
        $cell = $ws.Cells.Item($excelLig, $col + 1)
        $valeur = $ligne.($colonnes[$col])
        $cell.Value2 = $valeur
        $cell.Font.Size = 7




        function Get-ExcelColor($r, $g, $b) {
    return $r + ($g * 256) + ($b * 65536)
}

$palette = @(
    (Get-ExcelColor 183  222  232),  # Bleu
    (Get-ExcelColor 216  228  188),  # Vert
    (Get-ExcelColor 252 213   180),  # Orange
    (Get-ExcelColor 230  184  183),  # Rose foncé
    (Get-ExcelColor 204  192  218)   # Violet foncé
)

# Récupération des groupes
$groupes = Get-ADGroup -Filter {Name -like "GG_F_CDAD-BDX*"} | Sort-Object Name

# Récupération des utilisateurs
$utilisateurs = $groupes | ForEach-Object {
    Get-ADGroupMember -Identity $_ -Recursive | Where-Object { $_.objectClass -eq "user" }
} | Sort-Object SamAccountName | Select-Object -Unique -Property Name, SamAccountName

# Construction de la matrice
$matrice = foreach ($user in $utilisateurs) {
    $userAD = Get-ADUser $user.SamAccountName -Properties MemberOf, DisplayName
    $groupesUser = $userAD.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
    $ligne = [ordered]@{ "Utilisateur" = $userAD.SamAccountName }
    foreach ($groupe in $groupes) {
        $ligne[$groupe.Name] = if ($groupesUser -contains $groupe.Name) { "+" } else { "" }
    }
    [PSCustomObject]$ligne
}

# ---- EXCEL VIA COM ----
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = "Matrice Droits"

$colonnes = $matrice[0].PSObject.Properties.Name

# ---- EN-TÊTE ----
for ($col = 0; $col -lt $colonnes.Count; $col++) {
    $cell = $ws.Cells.Item(1, $col + 1)
    $cell.Value2 = $colonnes[$col]
    $cell.Font.Bold = $true
    $cell.Font.Size = 7

    if ($col -eq 0) {
        $cell.Interior.Color = Get-ExcelColor 255 255 255
        $cell.Font.Color = 0xFFFFFF
        $cell.VerticalAlignment = -4107
    } else {
        $cell.Interior.Color = $palette[($col - 1) % 5]
        $cell.Font.Color = 000000      # Texte blanc sur fond foncé
        $cell.Orientation = 90
        $cell.VerticalAlignment = -4107
        $cell.HorizontalAlignment = -4108
    }
}

# Hauteur ligne en-tête réduite au minimum viable avec rotation
$ws.Rows.Item(1).RowHeight = 100

# ---- DONNÉES ----
for ($lig = 0; $lig -lt $matrice.Count; $lig++) {
    $ligne = $matrice[$lig]
    $excelLig = $lig + 2

    for ($col = 0; $col -lt $colonnes.Count; $col++) {
        $cell = $ws.Cells.Item($excelLig, $col + 1)
        $valeur = $ligne.($colonnes[$col])
        $cell.Value2 = $valeur
        $cell.Font.Size = 7
        $cell.Borders.LineStyle = 1
        $cell.Borders.Weight = 1

        if ($col -eq 0) {
            if ($excelLig % 2 -eq 0) {
                $cell.Interior.Color = Get-ExcelColor 200 200 200
            } else {
                $cell.Interior.Color = Get-ExcelColor 235 235 235
            }
        } elseif ($valeur -eq "+") {
            $cell.Interior.Color = Get-ExcelColor 80 80 80
            $cell.Font.Color = 000000
            $cell.HorizontalAlignment = -4108
            $cell.Font.Bold = $true
        } else {
            if ($excelLig % 2 -eq 0) {
                $cell.Interior.Color = Get-ExcelColor 200 200 200
            } else {
                $cell.Interior.Color = Get-ExcelColor 235 235 235
            }
        }
    }

    # Hauteur de ligne minimale
    $ws.Rows.Item($excelLig).RowHeight = 11
}

# ---- FIGER VOLETS ----
$ws.Cells.Item(2, 2).Select()
$excel.ActiveWindow.FreezePanes = $true

# ---- LARGEUR COLONNES ----
$ws.Columns.Item(1).ColumnWidth = 15  # Colonne utilisateur
for ($col = 2; $col -le $colonnes.Count; $col++) {
    $ws.Columns.Item($col).ColumnWidth = 2.5  # Minimum lisible
}



$fichier = "C:\Users\x\Downloads\matrice_droits-test4.xlsx"
$wb.SaveAs($fichier)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Export terminé : $fichier" -ForegroundColor Green







        if ($col -eq 0) {
            if ($excelLig % 2 -eq 0) {
                $cell.Interior.Color = Get-ExcelColor 200 200 200
            } else {
                $cell.Interior.Color = Get-ExcelColor 235 235 235
            }
        } elseif ($valeur -eq "+") {
            $cell.Interior.Color = Get-ExcelColor 80 80 80
            $cell.Font.Color = 000000
            $cell.HorizontalAlignment = -4108
            $cell.Font.Bold = $true
        } else {
            if ($excelLig % 2 -eq 0) {
                $cell.Interior.Color = Get-ExcelColor 200 200 200
            } else {
                $cell.Interior.Color = Get-ExcelColor 235 235 235
            }
        }
    }

    # Hauteur de ligne minimale
    $ws.Rows.Item($excelLig).RowHeight = 11
}

# ---- FIGER VOLETS ----
$ws.Cells.Item(2, 2).Select()
$excel.ActiveWindow.FreezePanes = $true

# ---- LARGEUR COLONNES ----
$ws.Columns.Item(1).ColumnWidth = 15  # Colonne utilisateur
for ($col = 2; $col -le $colonnes.Count; $col++) {
    $ws.Columns.Item($col).ColumnWidth = 2.5  # Minimum lisible
}



$fichier = "C:\Users\x\Downloads\matrice_droits-test4.xlsx"
$wb.SaveAs($fichier)
$wb.Close()
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Export terminé : $fichier" -ForegroundColor Green
