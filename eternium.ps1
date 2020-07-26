function eternium($a,$b,$c,$d,$e){ 
$global:errpref = $ErrorActionPreference
	switch($a){
		open {
			$global:ie = new-object -ComObject "InternetExplorer.Application"
			$global:ie.visible = $true
			$global:ie.navigate($b)
            eternium busy
		}
		hide {
		$global:ie.visible = $false
		}
		show {
		$global:ie.visible = $true
		}
        busy {while($global:ie.Busy) { Start-Sleep -s 1 }
        while($global:ie.ReadyState() -ne 4){ Start-Sleep -s 1 }
        }
        navigate {
        if ($b){$global:ie.Navigate($b)
           eternium busy}           
           else{
           return $global:ie.LocationURL
           }
           }
		get {
                if ($b.toLower() -eq "innertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].innerText.toString())
                        if ($item -eq $c){
                            return $global:ie.document.getElementsByTagName('*')[$i]
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "innerhtml") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].innerhtml.toString())
                        if ($item -eq $c){
                            return $global:ie.document.getElementsByTagName('*')[$i]
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outerhtml") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].outerhtml.toString())
                        if ($item -eq $c){
                            return $global:ie.document.getElementsByTagName('*')[$i]
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].outertext.toString())
                        if ($item -eq $c){
                            return $global:ie.document.getElementsByTagName('*')[$i]
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                else {
                        $global:ie.document.getElementsByTagName('*') | % {
					    if ($_.getAttributeNode($b).Value -eq $c) {
						     return $_ 
					    }
				    }
                }
		    }
            compatget {
                    if ($b.toLower() -eq "innertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].innerText.toString())
                        if ($item -eq $c){
                            return $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i]
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "innerhtml") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].innerhtml.toString())
                        if ($item -eq $c){
                            return $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i]
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outerhtml") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].outerhtml.toString())
                        if ($item -eq $c){
                            return $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i]
                           
                        }
                        $i = $i+1
                    }
                     $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outertext") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].outertext.toString())
                        if ($item -eq $c){
                            return $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i]
                           
                        }
                        $i = $i+1
                    }
                     $ErrorActionPreference = $global:errpref
                }
                else {
                        $global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					    if ($_.getAttributeNode($b).Value -eq $c) {
						     return $_ 
					    }
				    }
                }
            }

            set {
                if ($b.toLower() -eq "innertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].innerText.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "innerhtml") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].innerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outerhtml") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].outerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].outertext.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                else {
                        $global:ie.document.getElementsByTagName('*') | % {
					    if ($_.getAttributeNode($b).Value -eq $c) {
						     $_.$d = $e
					    }
				    }
                }
		    }
            compatset {
                    if ($b.toLower() -eq "innertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].innerText.toString())
                        if ($item -eq $c){
                            $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "innerhtml") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].innerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outerhtml") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].outerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outertext") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].outertext.toString())
                        if ($item -eq $c){
                           $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].$d = $e
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                else {
                        $global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					    if ($_.getAttributeNode($b).Value -eq $c) {
						     $_.$d = $e
					    }
				    }
                }
            }

        
		oldset {
			try{
				$global:ie.document.getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						$_.$d = $e
					}
				}
			}
			catch{
	            $global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						 $_.$d = $e
					}
                } 
			}
        }

            click {
                if ($b.toLower() -eq "innertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].innerText.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].click()
                            eternium busy
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "innerhtml") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].innerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].click()
                            eternium busy
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outerhtml") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].outerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].click()
                            eternium busy
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.getElementsByTagName('*')[$i].outertext.toString())
                        if ($item -eq $c){
                            $global:ie.document.getElementsByTagName('*')[$i].click()
                            eternium busy
                           
                        }
                        $i = $i+1
                    }
                     $ErrorActionPreference = $global:errpref
                }
                else {
                $ErrorActionPreference = 'SilentlyContinue'
                        $global:ie.document.getElementsByTagName('*') | % {
					    if ($_.getAttributeNode($b).Value -eq $c) {
						     $_.click()
                            eternium busy
					    }
				    }
                    $ErrorActionPreference = $global:errpref
                }
		    }
            compatclick {
                    if ($b.toLower() -eq "innertext") {
                    $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].innerText.toString())
                        if ($item -eq $c){
                            $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].click()
                            eternium busy
                           
                        }
                        $i = $i+1
                    }
                     $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "innerhtml") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].innerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].click()
                            eternium busy
                           
                        }
                        $i = $i+1
                    }
                     $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outerhtml") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].outerhtml.toString())
                        if ($item -eq $c){
                            $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].click()
                            eternium busy
                            
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                elseif($b.toLower() -eq "outertext") {
                $ErrorActionPreference = 'SilentlyContinue'
				    $i = 0
                    while ($i -le $global:ie.document.IHTMLDocument3_getElementsByTagName('*').length()) {
                        $item = ($global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].outertext.toString())
                        if ($item -eq $c){
                           $global:ie.document.IHTMLDocument3_getElementsByTagName('*')[$i].click()
                            eternium busy
                        
                        }
                        $i = $i+1
                    }
                    $ErrorActionPreference = $global:errpref
                }
                else {
                        $global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					    if ($_.getAttributeNode($b).Value -eq $c) {
						    $_.click()
                            eternium busy
					    }
				    }
                }
            }

		oldclick {
			try{ 
				$global:ie.document.getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						$_.click()
                        eternium busy
					}
				}
			}
			catch{
				$global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						 $_.click()
                         eternium busy
					}
				}
			}
		}
	}
<#
    .SYNOPSIS
    Automates Internet Explorer.
     
    .DESCRIPTION
     VDS
	eternium open 'https://google.com'
	eternium navigate 'https://dialogshell.com'
	$currentpage = $(eternium navigate)
	$value = $(eternium get 'id' 'Text1').value
	$innertext = $(eternium get 'innerhtml' 'Hello<BR>There').innertext
	eternium set 'class' 'Text1' 'value' 'new value'
	eternium click 'name' 'button1'
		$value = $(eternium compatget 'id' 'Text1').value
		$innertext = $(eternium compatget 'innerhtml' 'Hello<BR>There').innertext
		eternium compatset 'class' 'Text1' 'value' 'new value'
		eternium compatclick 'name' 'button1'
	eternium hide
	eternium show
    
    .LINK
    https://dialogshell.com/vds/help/index.php/eternium
#>	
}
