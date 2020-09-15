$fileName = "C:\Users\Nick\OneDrive - Crickhowell High School\CHS.SIMSExchange\CHS.SIMSExchange\staffmapping.xml”;
$xmlDoc = [System.Xml.XmlDocument](Get-Content $fileName);
Get-ADUser -SearchBase "OU=Teaching Staff,OU=CRI,OU=Establishments,DC=crick,DC=internal" -Filter { Enabled -eq $True } -Properties Surname,GivenName,Mail,Enabled | Sort-Object Surname,GivenName |  ForEach-Object { 
    $s = $xmlDoc.CreateElement("staff");
    $f = $_.GivenName;
    $l = $_.Surname;
    $m = $_.Mail;
    if ($_.Enabled -eq $true -and $_.Mail -ne "" -and $m -ne $null) {
    $s.SetAttribute("first","$f"); 
    $s.SetAttribute("last","$l"); 
    $s.SetAttribute("email","$m"); 
    $xmlDoc.staffmappings.AppendChild($s);
    }
}

$xmlDoc.Save($fileName);