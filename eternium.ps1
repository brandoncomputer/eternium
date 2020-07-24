function eternium($a,$b,$c,$d,$e){ 
	switch($a){
		open {
			$global:ie = new-object -ComObject "InternetExplorer.Application"
			$global:ie.visible = $true
			$global:ie.navigate($b)
			while($global:ie.Busy) {
				Start-Sleep -Milliseconds 100
			}
		}
		get {
			try{ 
				$global:ie.document.getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						 return $_ 
					}
				}
			}
			catch{
				$global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						 return $_ 
					}
				}
			}
		}
		set {
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
			try{ 
				$global:ie.document.getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						$_.click()
					}
				}
			}
			catch{
				$global:ie.document.IHTMLDocument3_getElementsByTagName('*') | % {
					if ($_.getAttributeNode($b).Value -eq $c) {
						 $_.click()
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
	eternium open 'http://google.com'
	$value = $(eternium get 'id' 'Text1').value
	eternium set 'class' 'Text1' 'value' 'new value'
	eternium click 'name' 'button1'
    
    .LINK
    https://dialogshell.com/vds/help/index.php/eternium
#>	
}
