class Dictionary {
    [String[]]$Words
        Dictionary (){
    }
}

function txt_search(){
    $dict = New-Object Dictionary
    $dict.Words = @(Get-Content -Path .\languages\english.txt)
    $words = New-Object Dictionary
    $path = Read-Host -Prompt "Enter path to txt file" #example file.txt
    $words.Words = @(Get-Content -Path .\$path) -split ' '
    get_outcome $words $dict
    Read-Host -Prompt "Search completed. Push enter to return to main menu"
}

function xml_search(){
    $dict = New-Object Dictionary
    $words = New-Object Dictionary
    $dict.Words = @(Get-Content -Path .\languages\english.txt)
    $path = Read-Host -Prompt "Enter absolute path to xml file" #Example= C:\Users\username\OneDrive\Desktop\password_finder\computers.xml
    $xpath = Read-Host -Prompt "Enter xpath" #Example: Computers/Computer/Statement
    Select-Xml -Path $path -XPath $xpath | 
    ForEach-Object { 
        $words.Words = $_.Node.InnerXML -split '[\r\n ]'
        get_outcome $words $dict
    }
    Read-Host -Prompt "Search completed. Push enter to return to main menu"
}

function docx_search(){
    $dict = New-Object Dictionary
    $words = New-Object Dictionary
    $dict.Words = @(Get-Content -Path .\languages\english.txt)
    $source = Read-Host -Prompt "Enter absolute path of folder containing docx/rtf files/" #Example= C:\Users\username\OneDrive\Desktop\password_finder\
    $prefix = Read-Host -Prompt "Enter prefix"
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $False
    $docs = Get-ChildItem -Path $source | Where-Object {$_.Name -match $prefix} # Get all the files in the folder that ends with the prefix
    foreach ($name in $docs.Name){
        $Document = $Word.Documents.Open($source +'\' + $name)  
        foreach($paragraph in $Document.Paragraphs){
            $words.Words = $paragraph.Range.Text -split '[\r\n ]'
            get_outcome $words $dict
        }
        $Word.Application.ActiveDocument.Close()
    }
    Read-Host -Prompt "Search completed. Push enter to return to main menu"
}

function get_outcome($words, $dict){
    for($i=0; $i -le $words.Words.length; $i++){
        $match = 0
        for($j=0; $j -le $dict.Words.length; $j++){
            if($words.Words[$i] -eq $dict.Words[$j]){
                $match = 1
            }
        }
        if($match -eq 0 -and $words.Words[$i]-notmatch '[,.]'){ # important: Set regex filter
            $words.Words[$i]
        }
        
    }
}

function testing(){
    Expand-Archive -LiteralPath 'C:\Users\username\OneDrive\Desktop\password_finder\test.zip' -DestinationPath C:\Users\username\OneDrive\Desktop\password_finder\languages
}

function Main{

    $selection

    while($selection -notmatch '[q]'){
        Write-Host ""
        Write-Host "PowerShell Password Finder"
        Write-Host "--------------------------"
        Write-Host ""
        Write-Host "1. Find password in TXT file"
        Write-Host "2. Find password in XML file"
        Write-Host "3. Find password in DOCX/RTF files"
        Write-Host "4. Enter 'q' to quit"
        Write-Host "5. Enter 'x' to clear output"
        $selection = Read-Host -Prompt "Enter Selection"
        if($selection -eq "x"){
            Clear-Host
        }
        elseif($selection -notmatch '[1-7q]'){
            Read-Host -Prompt "Error. Input must be 1 to 4 or q. Press enter to continue"
        }
        elseif($selection -eq "q"){
            break
        }
        elseif($selection -eq 1){
            txt_search
        }
        elseif($selection -eq 2){
            xml_search
        }
        elseif($selection -eq 3){
            docx_search
        }
        elseif($selection -eq 6){
            testing
        }  
        elseif($selection -eq 7){
            Set-ConsoleOpacity 10
        }       
    }
    if($selection -eq 'q'){
        Read-Host -Prompt "Thank you for using PowerShell Password Finder. Goodbye."
    }
}

Main