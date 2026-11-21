# VBIDE.Import uses the machine ANSI codepage for file paths and content.
# Generate equivalent ASCII VBA source without changing business modules.
function ConvertTo-AsciiVba([string]$source){
    # VBA optional defaults must be compile-time literals. The presentation-only
    # default label is restored at entry; explicitly supplied empty labels stay empty.
    $source=$source.Replace('Optional ByVal label As String = "当前将处理"','Optional ByVal label As String = "__DH_DEFAULT_LABEL__"')
    $source=[regex]::Replace($source,'(Public Function DH_SelectionTarget[^\r\n]*\r?\n)','$1    If label = "__DH_DEFAULT_LABEL__" Then label = "当前将处理"'+"`r`n")
    $convert=[Text.RegularExpressions.MatchEvaluator]{param($match)
        $literal=$match.Value.Substring(1,$match.Length-2).Replace('""','"')
        if($literal -notmatch '[^\x00-\x7F]'){return $match.Value}
        $parts=[Collections.Generic.List[string]]::new()
        $buffer=''
        foreach($character in $literal.ToCharArray()){
            if([int]$character -lt 128){$buffer+=$character}else{
                if($buffer.Length){$parts.Add('"'+$buffer.Replace('"','""')+'"');$buffer=''}
                $code=[int]$character;if($code -gt 32767){$code-=65536}
                $parts.Add('ChrW('+ $code +')')
            }
        }
        if($buffer.Length){$parts.Add('"'+$buffer.Replace('"','""')+'"')}
        $groups=[Collections.Generic.List[string]]::new()
        for($i=0;$i -lt $parts.Count;$i+=8){$groups.Add(($parts[$i..([Math]::Min($i+7,$parts.Count-1))] -join ' & '))}
        return '('+($groups -join " & _`r`n        ")+')'
    }
    $lines=[Collections.Generic.List[string]]::new()
    foreach($line in ($source -split '\r?\n')){
        $quoted=$false;$comment=$line.Length
        for($i=0;$i -lt $line.Length;$i++){
            if($line[$i] -eq '"'){
                if($quoted -and $i+1 -lt $line.Length -and $line[$i+1] -eq '"'){$i++;continue}
                $quoted=-not $quoted
            }elseif(-not $quoted -and $line[$i] -eq "'"){$comment=$i;break}
        }
        $code=$line.Substring(0,$comment)
        $suffix=$line.Substring($comment)
        $code=[regex]::Replace($code,'"(?:""|[^"])*"',$convert)
        if($code -match '[^\x00-\x7F]'){throw 'Non-ASCII executable identifier cannot be converted safely'}
        $lines.Add($code+[regex]::Replace($suffix,'[^\x00-\x7F]','?'))
    }
    $projected=$lines -join "`r`n"
    $lineNumber=0;$continued=0
    foreach($line in ($projected -split '\r?\n')){
        $lineNumber++
        if($line.Length -gt 1023){throw ("VBA import physical line exceeds 1023 characters at projected line "+$lineNumber)}
        if($line -match '_\s*$'){$continued++}else{$continued=0}
        if($continued -gt 24){throw ("VBA import exceeds 24 continuations at projected line "+$lineNumber)}
    }
    return $projected
}
