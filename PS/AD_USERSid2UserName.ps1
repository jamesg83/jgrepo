Import-Module ActiveDirectory

$objSID = New-Object System.Security.Principal.SecurityIdentifier ` 
("S-1-5-21-1409082233-1644491937-839522115-11350") 
$objUser = $objSID.Translate( [System.Security.Principal.NTAccount]) 
$objUser.Value