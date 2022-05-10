cls

Get-Content -LiteralPath "$PSScriptRoot\config.txt" | Where-Object {$_ -like '$*'} | Invoke-Expression

$clneibs = @{}

function Run-MFBWalker ($startnode, $maxsteps)
{
    #оптимизация: когда кол-во шагов превышает лям, вести лог становится накладно - отключаем лог
    #$path_log = New-Object System.Collections.Generic.List[System.Object]
    $footprints = @{}

    for($step = 0; $step -lt $maxsteps; $step++)
    {
        #$path_log.Add( @($step,$script:strings[$startid]) )
        $footprints[$startnode]++

        $startnode = ($script:clneibs[$startnode] | Get-Random)
    }

    #$path_log | Out-GridView
    $footprints.GetEnumerator() | Out-GridView
}

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

$gml_str = "graph`r`n[`r`n`tdirected`t0`r`n"
$gml_str_children = New-Object System.Collections.Generic.List[System.Object]
$gml_str_neibs = New-Object System.Collections.Generic.List[System.Object]

#$stat = @{}
$stat = New-Object System.Collections.Generic.List[System.Object]

foreach($node in $nodes)
{
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $richTextBox1.LoadFile($node.FullName)

    $meat = ($node.Length - 155)/7100.0
    $red = [int](248 - (248 - 164)*$meat)
    $green = [int](221 - (221 - 99)*$meat)
    $blue = [int](177 - (177 - 18)*$meat)

    if($red -lt 0) {$red = 0}
    if($green -lt 0) {$green = 0}
    if($blue -lt 0) {$blue = 0}

    $gml_color = '"#' + ('{0:X}{1:X}{2:X}' -f $red,$green,$blue) + '"'

    $gml_str_children.Add("`t" + 'node' + "`r`n" + `
         "`t" + "[" + "`r`n" + `
         "`t`t" + "id`t" + $node2id[$node.BaseName] + "`r`n" + `
         "`t`t" + "label`t" + '"' + `
         ($node.BaseName).Replace('"',"'").Replace('[',"_").Replace(']',"_").Replace('\',"_").Replace('/',"_").Replace(',',"_").Replace('&',"_") + `
         '"' + "`r`n" + `
         "`t`t" + "graphics" + "`r`n" + `
         "`t`t" + "[" + "`r`n" + `
         "`t`t`t" + "x`t" + (Get-Random -Minimum 0 -Maximum 10000) + "`r`n" + `
         "`t`t`t" + "y`t" + (Get-Random -Minimum 0 -Maximum 10000) + "`r`n" + `
         "`t`t`t" + "fill`t" + $gml_color + "`r`n" + `
         "`t`t" + "]" + "`r`n" + `
         "`t" + "]" + "`r`n")

    $links = ($richTextBox1.Lines | Where-Object {$_ -like '*linkto:*'} | ForEach-Object {$_.Split(':')[1]})
    #$links

    $clneibs[$node.BaseName] = @() + ( $links | Where-Object {$node2id[$_]} )

    #$stat[$node.BaseName] = $links.Count

    $fobj = Get-Item -LiteralPath $node.FullName

    $stat.Add([pscustomobject]@{
        'Node' = $node.BaseName
        'Links' = $links.Count
        'Created' = $fobj.CreationTime
        'Modyfied' = $fobj.LastWriteTime
        'Size' = $fobj.Length
        })

    foreach($link in $links)
    {
        if(-not $node2id[$link])
        {
            continue
        }

        if($node2id[$node.BaseName] -gt $node2id[$link])
        {
            continue
        }
            
        $gml_str_neibs.Add("`t" + 'edge' + "`r`n" + `
        "`t" + "[" + "`r`n" + `
        "`t`t" + "source`t" + $node2id[$node.BaseName] + "`r`n" + `
        "`t`t" + "target`t" + $node2id[$link] + "`r`n" + `
        "`t" + "]" + "`r`n")
    }
}

$gml_str += (-join $gml_str_children)
$gml_str += (-join $gml_str_neibs)
$gml_str += "]`r`n"

$gml_str | Out-File -FilePath $out_gml -Encoding default

$stat | Out-GridView

#Run-MFBWalker '!!!!!_start' 1000000