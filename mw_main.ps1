﻿cls

$workfilearg = $script:args

Get-Content -LiteralPath "$PSScriptRoot\config.txt" | Where-Object {$_ -like '$*'} | Invoke-Expression

$antibonus = @{}
$review_ts = @{}

function Add-Node ($RootNode,$NodeName)
{
    $newNode = new-object System.Windows.Forms.TreeNode
    $newNode.Name = $NodeName
    $newNode.Text = $NodeName
    $newNode.Tag = $NodeName
    $Null = $RootNode.Nodes.Add($newNode)
    return $newNode
}


function Get-NeighborLinks ($node)
{
    $nodefilepath = "$($script:workdirectory)\$($node).dcmp2"

    if(-not (Test-Path -LiteralPath $nodefilepath))
    {
        return @()
    }

    $tmp_richTextBox = New-Object System.Windows.Forms.RichTextBox
    $tmp_richTextBox.LoadFile($nodefilepath)
    
    $links = ($tmp_richTextBox.Lines | Where-Object {$_ -like '*linkto:*'} | ForEach-Object {$_.Split(':')[1]})
    $links2 = @()
    
    $word2links = ($tmp_richTextBox.Lines | Where-Object {$_ -like '*word2links:*'} | ForEach-Object {$_.Split(':')[1]})
    
    foreach($word2link in $word2links)
    {
        $links2 += (Search-Word $word2link).Node
    }
    
    $links = @() + $links + ($links2 | Sort-Object | Get-Unique)

    return $links
}


function GoToPage ($fname)
{
    $script:workfile = Get-ChildItem -LiteralPath $fname
    $script:workdirectory = $script:workfile.Directory.FullName
    $script:workfilebasename = $script:workfile.BaseName
    $script:workfilefullname = $script:workfile.FullName

    $form1.Text = "MyWiki - $($script:workfilefullname)"
    $richTextBox1.LoadFile($script:workfilefullname)
    $label1.Text = ""

    $treeView1.Nodes.Clear()
    Add-Node $treeView1 $script:workfilebasename | Out-Null
}


function Search-Word ($request)
{
    $result = New-Object System.Collections.Generic.List[System.Object]

    $nodes = Get-ChildItem -LiteralPath $script:workdirectory -Recurse '*.dcmp2'

    $tmp_richTextBox = New-Object System.Windows.Forms.RichTextBox

    foreach($node in $nodes)
    {
        $tmp_richTextBox.LoadFile($node.FullName)
        $matchstrings = ($tmp_richTextBox.Lines | Where-Object {$_ -like "*$($request)*"})

        foreach($mstr in $matchstrings)
        {
            $result.Add([pscustomobject]@{
            'Node' = $node.BaseName
            'Line' = $mstr
            })
        }
    }

    return $result
}


function Review-Next
{
    $results = New-Object System.Collections.Generic.List[System.Object]

    $nodes = Get-ChildItem -LiteralPath $script:workdirectory -Recurse '*.dcmp2'

    $tmp_richTextBox = New-Object System.Windows.Forms.RichTextBox

    foreach($node in $nodes)
    {
        $tmp_richTextBox.LoadFile($node.FullName)
        $matchstrings = ($tmp_richTextBox.Lines | Where-Object {$_ -like '#review:*'})

        $cand_ts = $node.LastWriteTime

        foreach($mstr in $matchstrings)
        {
            $tmp_arr = $mstr.Split(':')
            $tmp_arr = ($tmp_arr[1]).Split(' ')
            $cand_rev = [int]($tmp_arr[0])

            $degree = 250/([math]::Pow($cand_rev,1.7))

            if(-not $script:review_ts.ContainsKey($node.BaseName))
            {
                $script:review_ts[$node.BaseName] = $cand_ts
            }

            if(((Get-Date) - $script:review_ts[$node.BaseName]).TotalMinutes -lt $script:antibonus[$node.BaseName]*3*$cand_rev)
            {
                continue
            }
            
            for($dd = 0; $dd -lt $degree; $dd++)
            {
                $results.Add([pscustomobject]@{
                'NodeName' = $node.BaseName
                'RevDays' = $cand_rev
                'Updated' = $cand_ts
                'instance' = $dd
                'degree' = $degree
                'rev_ts' = $script:review_ts[$node.BaseName]
                })
            }
        }
    }

    if($results.Count -eq 0)
    {
        Write-Host "Нет больше заметок на ревью"
        return
    }

    Write-Host "Конкуренция: $($results.Count)"

    ### $selected_link = ($results.GetEnumerator() | Sort-Object -Property Updated | Out-GridView -OutputMode Single).NodeName
    $selected_link = ($results.GetEnumerator() | Get-Random).NodeName
    
    $script:antibonus[$selected_link]++
    $script:review_ts[$selected_link] = (Get-Date)

    $target = "$($script:workdirectory)\$($selected_link).dcmp2"
    Write-Host "Go for review $target"
    GoToPage $target

    ## debug
    # $results.GetEnumerator() | Out-GridView
}


#Generated Form Function
function GenerateForm {
########################################################################
# Code Generated By: SAPIEN Technologies PrimalForms (Community Edition) v1.0.10.0
# Generated On: 24.11.2021 19:13
# Generated By: vlad
########################################################################

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$treeView1 = New-Object System.Windows.Forms.TreeView
$ReviewButton = New-Object System.Windows.Forms.Button
$RecentButton = New-Object System.Windows.Forms.Button
$ConsistencyCheckerButton = New-Object System.Windows.Forms.Button
$GitPushButton = New-Object System.Windows.Forms.Button
$GitPullButton = New-Object System.Windows.Forms.Button
$GitStatusButton = New-Object System.Windows.Forms.Button
$FormatClearButton = New-Object System.Windows.Forms.Button
$FormatBlueButton = New-Object System.Windows.Forms.Button
$FormatGreenButton = New-Object System.Windows.Forms.Button
$FormatRedButton = New-Object System.Windows.Forms.Button
$SaveButton = New-Object System.Windows.Forms.Button
$AddLinkToExistPageButton = New-Object System.Windows.Forms.Button
$SearchButton = New-Object System.Windows.Forms.Button
$JumpPageButton = New-Object System.Windows.Forms.Button
$GoToPageButton = New-Object System.Windows.Forms.Button
$label1 = New-Object System.Windows.Forms.Label
$richTextBox1 = New-Object System.Windows.Forms.RichTextBox
$FormatCodeButton = New-Object System.Windows.Forms.Button
$FormatUnderlineButton = New-Object System.Windows.Forms.Button
$FormatItalicButton = New-Object System.Windows.Forms.Button
$FormatBoldButton = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
#Provide Custom Code for events specified in PrimalForms.
$handler_treeView1_AfterSelect= 
{
    $links = (Get-NeighborLinks $_.Node.Name)

    $_.Node.Nodes.Clear()

    foreach($link in $links)
    {
        Add-Node $_.Node $link
    }
}

$handler_treeView1_AfterExpand= 
{
    #Write-Host "EXPAND!" $_.Node.Name
}

$ReviewButton_OnClick= 
{
    Write-Host '----- Review ------------------------------------------------------'
    Review-Next
    Write-Host '-------------------------------------------------------------------'
}

$RecentButton_OnClick= 
{
    $rcnt_ts = (get-date).AddHours(-4)
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $request = [Microsoft.VisualBasic.Interaction]::InputBox("Введите временную отметку", "Недавние", $rcnt_ts.ToString("yyyy-MM-dd HH:mm"))
    $rcnt_ts = [datetime]::parseexact($request, "yyyy-MM-dd HH:mm", $null)
    #Write-Host $rcnt_ts

    $rcnt_nodes = (Get-ChildItem -LiteralPath $script:workdirectory -Recurse '*.dcmp2' | Where-Object {$_.LastWriteTime -ge $rcnt_ts} | Sort-Object -Property LastWriteTime -Descending).BaseName
    
    $treeView1.Nodes.Clear()

    foreach($rcnt_node in $rcnt_nodes)
    {
        Add-Node $treeView1 $rcnt_node | Out-Null
    }
}

$ConsistencyCheckerButton_OnClick= 
{
    Write-Host '----- Consistency Check -------------------------------------------'
    & "$PSScriptRoot\Consistency_Checker.ps1" "$($script:workdirectory)"
    Write-Host '-------------------------------------------------------------------'
}

$GitStatusButton_OnClick= 
{
    Write-Host '----- Git status --------------------------------------------------'
    #Start-Process -FilePath $script:path2git -ArgumentList 'status' -Wait -WorkingDirectory $script:workfile.Directory.Parent.FullName -NoNewWindow

    #https://stackoverflow.com/questions/8761888/capturing-standard-out-and-error-with-start-process/33652732#33652732

    $pinfo = New-Object System.Diagnostics.ProcessStartInfo
    $pinfo.FileName = $script:path2git
    $pinfo.WorkingDirectory = $script:workfile.Directory.Parent.FullName
    $pinfo.RedirectStandardError = $true
    $pinfo.RedirectStandardOutput = $true
    $pinfo.UseShellExecute = $false
    $pinfo.Arguments = 'status'
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pinfo
    $p.Start() | Out-Null
    $p.WaitForExit()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    #Write-Host "stdout: $stdout"
    #Write-Host "stderr: $stderr"
    #Write-Host "exit code: " + $p.ExitCode

    if(($stdout.Substring(0,14) -eq 'On branch main') -and
       ($stdout.Substring(15,30) -eq 'Your branch is up to date with') -and
       ($stdout.Substring(62,37) -eq 'nothing to commit, working tree clean'))
    {
        Write-Host 'OK!'
    }
    else
    {
        Write-Host "stdout: $stdout"
    }

    Write-Host '-------------------------------------------------------------------'
}

$GitPullButton_OnClick= 
{
    Write-Host '----- Git PULL ----------------------------------------------------'
    Start-Process -FilePath $script:path2git -ArgumentList 'pull' -Wait -WorkingDirectory $script:workfile.Directory.Parent.FullName -NoNewWindow
    Write-Host '-------------------------------------------------------------------'
}

$GitPushButton_OnClick= 
{
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $commit_message = [Microsoft.VisualBasic.Interaction]::InputBox("Введите комментарий к коммиту", "Коммит")

    Write-Host '----- Add, commit & push ------------------------------------------'
    Write-Host ">> git add -A`n`n"

    Start-Process -FilePath $script:path2git -ArgumentList 'add -A' -Wait -WorkingDirectory $script:workfile.Directory.Parent.FullName -NoNewWindow
    Start-Sleep -Seconds 3

    Write-Host ">> git commit -m `"$commit_message`"`n`n"

    Start-Process -FilePath $script:path2git -ArgumentList "commit -m `"$commit_message`"" -Wait -WorkingDirectory $script:workfile.Directory.Parent.FullName -NoNewWindow
    Start-Sleep -Seconds 3

    Write-Host ">> git push`n`n"

    Start-Process -FilePath $script:path2git -ArgumentList 'push' -Wait -WorkingDirectory $script:workfile.Directory.Parent.FullName -NoNewWindow
    Start-Sleep -Seconds 3

    Write-Host '-------------------------------------------------------------------'
}

$FormatRedButton_OnClick= 
{
    $richTextBox1.SelectionColor = [System.Drawing.Color]::FromArgb(255,255,0,0)
}

$FormatGreenButton_OnClick= 
{
    $richTextBox1.SelectionColor = [System.Drawing.Color]::FromArgb(255,0,128,0)
}

$FormatBlueButton_OnClick= 
{
    $richTextBox1.SelectionColor = [System.Drawing.Color]::FromArgb(255,0,0,205)
}

$FormatClearButton_OnClick= 
{
    #clear formatting
    $richTextBox1.SelectionFont = New-Object System.Drawing.Font("Tahoma",10,0,3,0) #Microsoft Sans Serif
    $richTextBox1.SelectionColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
}

$SaveButton_OnClick= 
{
    #save changes
    $richTextBox1.SaveFile($script:workfilefullname)
    $label1.Text = "СОХРАНЕНО!"
}

$FormatCodeButton_OnClick= 
{
    $richTextBox1.SelectionFont = New-Object System.Drawing.Font("Lucida Console",8,0,3,204)
}

$SearchButton_OnClick= 
{
    #поиск текста
    
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $request = [Microsoft.VisualBasic.Interaction]::InputBox("Введите искомое слово/фрагмент", "Поиск")

    $searchresults = (Search-Word $request)

    $searchresults | Out-GridView
}

$RTFChanged= 
{
    $label1.Text = "Файл $($script:workfilefullname) был изменён"
}

$AddLinkToExistPageButton_OnClick= 
{
    #Add link
    $selected_link = ((Get-ChildItem -LiteralPath $script:workdirectory -Recurse '*.dcmp2').BaseName | Out-GridView -OutputMode Single)
    $richTextBox1.AppendText("`nlinkto:$selected_link")

    #вставляем обратную ссылку с помощью лайфхака
    $TempInvisibleRichTextBox = New-Object System.Windows.Forms.RichTextBox
    $TempInvisibleRichTextBox.LoadFile("$($script:workdirectory)\$($selected_link).dcmp2")
    $TempInvisibleRichTextBox.AppendText("`nНа эту страницу ссылается - linkto:$($script:workfilebasename)")
    $TempInvisibleRichTextBox.SaveFile("$($script:workdirectory)\$($selected_link).dcmp2")
}

$FormatBoldButton_OnClick= 
{
    #text to bold
    $sel_start = $richTextBox1.SelectionStart
    $sel_len = $richTextBox1.SelectionLength
    $sel_font = $richTextBox1.SelectionFont
    $richTextBox1.SelectionFont = New-Object System.Drawing.Font($sel_font.Name,$sel_font.SizeInPoints,($sel_font.Style -bxor [System.Drawing.FontStyle]::Bold))
    #$richTextBox1.Select($sel_start,$sel_len)
}

$GoToPageButton_OnClick= 
{
    #GoTo
    <#
    $links = (Get-NeighborLinks $script:workfilebasename)

    if($links.Count -eq 0)
    {
        Write-Host "No links"
        return #https://ridicurious.com/2020/01/23/deep-dive-break-continue-return-exit-in-powershell/
    }

    $selected_link = ($links | Out-GridView -OutputMode Single)
    #>
    $selected_link = $treeView1.SelectedNode.Name

    if(-not $selected_link)
    {
        Write-Host "No link selected"
        return
    }

    $prev = $script:workfilebasename
    $target = "$($script:workdirectory)\$($selected_link).dcmp2"

    if(Test-Path -LiteralPath $target)
    {
        Write-Host "Go to existing $target"
        GoToPage $target
    }
    else
    {
        Write-Host "Create $target"
        Copy-Item -LiteralPath "$($script:workdirectory)\template_blank.dcmp2" -Destination $target
        GoToPage $target
        $richTextBox1.AppendText("$selected_link")
        $richTextBox1.AppendText("`n#review:1")
        $richTextBox1.AppendText("`n`nОсновной текст")
        $richTextBox1.AppendText("`n`nНазад - linkto:$prev")
        $richTextBox1.SaveFile($target)
        GoToPage $target
    }
}

$FormatItalicButton_OnClick= 
{
    #text to italic
    $sel_start = $richTextBox1.SelectionStart
    $sel_len = $richTextBox1.SelectionLength
    $sel_font = $richTextBox1.SelectionFont
    $richTextBox1.SelectionFont = New-Object System.Drawing.Font($sel_font.Name,$sel_font.SizeInPoints,($sel_font.Style -bxor [System.Drawing.FontStyle]::Italic))
    #$richTextBox1.Select($sel_start,$sel_len)
}

$FormatUnderlineButton_OnClick= 
{
    #text to underline
    $sel_start = $richTextBox1.SelectionStart
    $sel_len = $richTextBox1.SelectionLength
    $sel_font = $richTextBox1.SelectionFont
    $richTextBox1.SelectionFont = New-Object System.Drawing.Font($sel_font.Name,$sel_font.SizeInPoints,($sel_font.Style -bxor [System.Drawing.FontStyle]::Underline))
}

$JumpPageButton_OnClick= 
{
    #Jump
    $selected_link = ((Get-ChildItem -LiteralPath $script:workdirectory -Recurse '*.dcmp2').FullName | Out-GridView -OutputMode Single)
    GoToPage $selected_link
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 762
$System_Drawing_Size.Width = 1192
$form1.ClientSize = $System_Drawing_Size
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$form1.Name = "form1"
$form1.Text = "Primal Form"


$treeView1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 43
$treeView1.Location = $System_Drawing_Point
$treeView1.Name = "treeView1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 680
$System_Drawing_Size.Width = 424
$treeView1.Size = $System_Drawing_Size
$treeView1.TabIndex = 21
$treeView1.add_AfterSelect($handler_treeView1_AfterSelect)
$treeView1.add_AfterExpand($handler_treeView1_AfterExpand)

$form1.Controls.Add($treeView1)


$ReviewButton.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 462
$ReviewButton.Location = $System_Drawing_Point
$ReviewButton.Name = "ReviewButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$ReviewButton.Size = $System_Drawing_Size
$ReviewButton.TabIndex = 20
$ReviewButton.Text = "Review"
$ReviewButton.UseVisualStyleBackColor = $True
$ReviewButton.add_Click($ReviewButton_OnClick)

$form1.Controls.Add($ReviewButton)


$RecentButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 432
$RecentButton.Location = $System_Drawing_Point
$RecentButton.Name = "RecentButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$RecentButton.Size = $System_Drawing_Size
$RecentButton.TabIndex = 19
$RecentButton.Text = "Recent"
$RecentButton.UseVisualStyleBackColor = $True
$RecentButton.add_Click($RecentButton_OnClick)

$form1.Controls.Add($RecentButton)


$ConsistencyCheckerButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 402
$ConsistencyCheckerButton.Location = $System_Drawing_Point
$ConsistencyCheckerButton.Name = "ConsistencyCheckerButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$ConsistencyCheckerButton.Size = $System_Drawing_Size
$ConsistencyCheckerButton.TabIndex = 18
$ConsistencyCheckerButton.Text = "Consistency check"
$ConsistencyCheckerButton.UseVisualStyleBackColor = $True
$ConsistencyCheckerButton.add_Click($ConsistencyCheckerButton_OnClick)

$form1.Controls.Add($ConsistencyCheckerButton)


$GitPushButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 341
$GitPushButton.Location = $System_Drawing_Point
$GitPushButton.Name = "GitPushButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$GitPushButton.Size = $System_Drawing_Size
$GitPushButton.TabIndex = 17
$GitPushButton.Text = "Git push"
$GitPushButton.UseVisualStyleBackColor = $True
$GitPushButton.add_Click($GitPushButton_OnClick)

$form1.Controls.Add($GitPushButton)


$GitPullButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 311
$GitPullButton.Location = $System_Drawing_Point
$GitPullButton.Name = "GitPullButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$GitPullButton.Size = $System_Drawing_Size
$GitPullButton.TabIndex = 16
$GitPullButton.Text = "Git pull"
$GitPullButton.UseVisualStyleBackColor = $True
$GitPullButton.add_Click($GitPullButton_OnClick)

$form1.Controls.Add($GitPullButton)


$GitStatusButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 281
$GitStatusButton.Location = $System_Drawing_Point
$GitStatusButton.Name = "GitStatusButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$GitStatusButton.Size = $System_Drawing_Size
$GitStatusButton.TabIndex = 15
$GitStatusButton.Text = "Git status"
$GitStatusButton.UseVisualStyleBackColor = $True
$GitStatusButton.add_Click($GitStatusButton_OnClick)

$form1.Controls.Add($GitStatusButton)


$FormatClearButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 685
$System_Drawing_Point.Y = 13
$FormatClearButton.Location = $System_Drawing_Point
$FormatClearButton.Name = "FormatClearButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 75
$FormatClearButton.Size = $System_Drawing_Size
$FormatClearButton.TabIndex = 14
$FormatClearButton.Text = "Clear"
$FormatClearButton.UseVisualStyleBackColor = $True
$FormatClearButton.add_Click($FormatClearButton_OnClick)

$form1.Controls.Add($FormatClearButton)


$FormatBlueButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatBlueButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)
$FormatBlueButton.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,205)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 655
$System_Drawing_Point.Y = 13
$FormatBlueButton.Location = $System_Drawing_Point
$FormatBlueButton.Name = "FormatBlueButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 23
$FormatBlueButton.Size = $System_Drawing_Size
$FormatBlueButton.TabIndex = 13
$FormatBlueButton.Text = "B"
$FormatBlueButton.UseVisualStyleBackColor = $True
$FormatBlueButton.add_Click($FormatBlueButton_OnClick)

$form1.Controls.Add($FormatBlueButton)


$FormatGreenButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatGreenButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)
$FormatGreenButton.ForeColor = [System.Drawing.Color]::FromArgb(255,0,128,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 625
$System_Drawing_Point.Y = 13
$FormatGreenButton.Location = $System_Drawing_Point
$FormatGreenButton.Name = "FormatGreenButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 23
$FormatGreenButton.Size = $System_Drawing_Size
$FormatGreenButton.TabIndex = 12
$FormatGreenButton.Text = "G"
$FormatGreenButton.UseVisualStyleBackColor = $True
$FormatGreenButton.add_Click($FormatGreenButton_OnClick)

$form1.Controls.Add($FormatGreenButton)


$FormatRedButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatRedButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)
$FormatRedButton.ForeColor = [System.Drawing.Color]::FromArgb(255,255,0,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 595
$System_Drawing_Point.Y = 13
$FormatRedButton.Location = $System_Drawing_Point
$FormatRedButton.Name = "FormatRedButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 23
$FormatRedButton.Size = $System_Drawing_Size
$FormatRedButton.TabIndex = 11
$FormatRedButton.Text = "R"
$FormatRedButton.UseVisualStyleBackColor = $True
$FormatRedButton.add_Click($FormatRedButton_OnClick)

$form1.Controls.Add($FormatRedButton)


$SaveButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 223
$SaveButton.Location = $System_Drawing_Point
$SaveButton.Name = "SaveButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$SaveButton.Size = $System_Drawing_Size
$SaveButton.TabIndex = 10
$SaveButton.Text = "Save changes"
$SaveButton.UseVisualStyleBackColor = $True
$SaveButton.add_Click($SaveButton_OnClick)

$form1.Controls.Add($SaveButton)


$AddLinkToExistPageButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 163
$AddLinkToExistPageButton.Location = $System_Drawing_Point
$AddLinkToExistPageButton.Name = "AddLinkToExistPageButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$AddLinkToExistPageButton.Size = $System_Drawing_Size
$AddLinkToExistPageButton.TabIndex = 9
$AddLinkToExistPageButton.Text = "Add link"
$AddLinkToExistPageButton.UseVisualStyleBackColor = $True
$AddLinkToExistPageButton.add_Click($AddLinkToExistPageButton_OnClick)

$form1.Controls.Add($AddLinkToExistPageButton)


$SearchButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 133
$SearchButton.Location = $System_Drawing_Point
$SearchButton.Name = "SearchButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$SearchButton.Size = $System_Drawing_Size
$SearchButton.TabIndex = 8
$SearchButton.Text = "Search"
$SearchButton.UseVisualStyleBackColor = $True
$SearchButton.add_Click($SearchButton_OnClick)

$form1.Controls.Add($SearchButton)


$JumpPageButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 73
$JumpPageButton.Location = $System_Drawing_Point
$JumpPageButton.Name = "JumpPageButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$JumpPageButton.Size = $System_Drawing_Size
$JumpPageButton.TabIndex = 7
$JumpPageButton.Text = "Jump"
$JumpPageButton.UseVisualStyleBackColor = $True
$JumpPageButton.add_Click($JumpPageButton_OnClick)

$form1.Controls.Add($JumpPageButton)


$GoToPageButton.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1071
$System_Drawing_Point.Y = 43
$GoToPageButton.Location = $System_Drawing_Point
$GoToPageButton.Name = "GoToPageButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 109
$GoToPageButton.Size = $System_Drawing_Size
$GoToPageButton.TabIndex = 6
$GoToPageButton.Text = "GoTo"
$GoToPageButton.UseVisualStyleBackColor = $True
$GoToPageButton.add_Click($GoToPageButton_OnClick)

$form1.Controls.Add($GoToPageButton)

$label1.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 443
$System_Drawing_Point.Y = 730
$label1.Location = $System_Drawing_Point
$label1.Name = "label1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 621
$label1.Size = $System_Drawing_Size
$label1.TabIndex = 5
$label1.Text = "label1"

$form1.Controls.Add($label1)

$richTextBox1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 443
$System_Drawing_Point.Y = 43
$richTextBox1.Location = $System_Drawing_Point
$richTextBox1.Name = "richTextBox1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 680
$System_Drawing_Size.Width = 621
$richTextBox1.Size = $System_Drawing_Size
$richTextBox1.TabIndex = 4
$richTextBox1.Text = ""
$richTextBox1.add_TextChanged($RTFChanged)

$form1.Controls.Add($richTextBox1)


$FormatCodeButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatCodeButton.Font = New-Object System.Drawing.Font("Lucida Console",8,0,3,204)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 533
$System_Drawing_Point.Y = 13
$FormatCodeButton.Location = $System_Drawing_Point
$FormatCodeButton.Name = "FormatCodeButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 55
$FormatCodeButton.Size = $System_Drawing_Size
$FormatCodeButton.TabIndex = 3
$FormatCodeButton.Text = "code"
$FormatCodeButton.UseVisualStyleBackColor = $True
$FormatCodeButton.add_Click($FormatCodeButton_OnClick)

$form1.Controls.Add($FormatCodeButton)


$FormatUnderlineButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatUnderlineButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,4,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 503
$System_Drawing_Point.Y = 13
$FormatUnderlineButton.Location = $System_Drawing_Point
$FormatUnderlineButton.Name = "FormatUnderlineButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 23
$FormatUnderlineButton.Size = $System_Drawing_Size
$FormatUnderlineButton.TabIndex = 2
$FormatUnderlineButton.Text = "U"
$FormatUnderlineButton.UseVisualStyleBackColor = $True
$FormatUnderlineButton.add_Click($FormatUnderlineButton_OnClick)

$form1.Controls.Add($FormatUnderlineButton)


$FormatItalicButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatItalicButton.Font = New-Object System.Drawing.Font("Courier New",8.25,2,3,204)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 473
$System_Drawing_Point.Y = 13
$FormatItalicButton.Location = $System_Drawing_Point
$FormatItalicButton.Name = "FormatItalicButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 23
$FormatItalicButton.Size = $System_Drawing_Size
$FormatItalicButton.TabIndex = 1
$FormatItalicButton.Text = "i"
$FormatItalicButton.UseVisualStyleBackColor = $True
$FormatItalicButton.add_Click($FormatItalicButton_OnClick)

$form1.Controls.Add($FormatItalicButton)


$FormatBoldButton.DataBindings.DefaultDataSourceUpdateMode = 0
$FormatBoldButton.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 443
$System_Drawing_Point.Y = 13
$FormatBoldButton.Location = $System_Drawing_Point
$FormatBoldButton.Name = "FormatBoldButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 23
$System_Drawing_Size.Width = 23
$FormatBoldButton.Size = $System_Drawing_Size
$FormatBoldButton.TabIndex = 0
$FormatBoldButton.Text = "B"
$FormatBoldButton.UseVisualStyleBackColor = $True
$FormatBoldButton.add_Click($FormatBoldButton_OnClick)

$form1.Controls.Add($FormatBoldButton)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form

GoToPage $workfilearg

$form1.ShowDialog()| Out-Null

} #End Function

#Call the Function
GenerateForm
