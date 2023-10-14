#cls

Write-Host $args[0]

#Get-Content -LiteralPath "$PSScriptRoot\config.txt" | Where-Object {$_ -like '$*'} | Invoke-Expression
$input_dir = $args[0]

[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null

$nodes = Get-ChildItem -LiteralPath $input_dir -Recurse '*.dcmp2'
#$nodes

$curr_id = 1
$node2id = @{}

foreach($node in $nodes)
{
    $node2id[$node.BaseName] = $curr_id
    $curr_id++
}

$neibs = @{}

$warn_count = 0

foreach($node in $nodes)
{
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $richTextBox1.LoadFile($node.FullName)
    $links = ($richTextBox1.Lines | Where-Object {$_ -like '*linkto:*'} | ForEach-Object {$_.Split(':')[1]})

    $tmplist = New-Object System.Collections.Generic.List[System.Object]

    foreach($link in $links)
    {
        if(-not $node2id[$link])
        {
            $warn_count++
            Write-Host "ПРЕДУПРЕЖДЕНИЕ: ссылка на несуществующую страницу: $link"
            continue
        }

        $tmplist.Add($node2id[$link])
    }

    $neibs[$node2id[$node.BaseName]] = $tmplist
}

#$neibs.GetEnumerator() | Out-GridView
write-host "Количество узлов: " $neibs.Count

$kosyak_count = 0

foreach($Aid in $neibs.Keys)
{
    #Write-Host $Aid
    foreach($Bid in $neibs[$Aid])
    {
        if(-not ($neibs[$Bid] -contains $Aid))
        {
            $Aname = ($node2id.GetEnumerator() | Where-Object {$_.value -eq $Aid}).Name
            $Bname = ($node2id.GetEnumerator() | Where-Object {$_.value -eq $Bid}).Name
            Write-Host "ОШИБКА: отсутствует ссылка: $Bname -> $Aname"
            $kosyak_count++
        }
    }
}

Write-Host "Ошибки:         $kosyak_count"
Write-Host "Предупреждения: $warn_count"