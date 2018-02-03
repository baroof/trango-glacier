<#
.Synopsis
   Report of OUs that contain nothing but GPOs, and fully empty OUs.
.DESCRIPTION
   "Empty" as in no users, computers, or other ADObjects were found inside the OU.
   The report also reformats the DistinguishedName to be more human-readable.
   Any listed OUs are potentially old, unused, and/or ripe for deletion.
.EXAMPLE
   .\Get-EmptyOUsAndTheirGPOs.ps1

.NOTES
  Version:        1.0
  Author:         Will Mooreston
  Creation Date:  2017-08-08
#>


# catch'm all
$empty_org_units = Get-ADOrganizationalUnit -Filter * | 
Where-Object {-not ( Get-ADObject -Filter * -SearchBase $_.Distinguishedname -SearchScope OneLevel -ResultSetSize 1 )}

# initialize the big hash
$ou_gpo_hash = @{} 

# cycle through each OU and find any linked GPOs
foreach($org_unit in $empty_org_units){
    
    # convert OU distinguished name to the AD folder path for easier locating
    $ou_path = ($org_unit | Select-Object -ExpandProperty DistinguishedName).ToString() -split ',' -replace 'DC=\w+' -replace 'OU='
    $ou_path = $ou_path[$ou_path.count..0] -join '\' -replace '^\\'

    # simultaneously add each OU to the hash while initializing the array for linked GPOs
    $ou_gpo_hash.$ou_path = @()
    
    # gather the GPO links
    $gpo_links = $org_unit | Select-Object -expand distinguishedname | Get-GPInheritance | Select-Object -expand gpolinks
    
    # append each link to the list for each OU w/in the hash
    foreach ($link in $gpo_links) {
        $gpo = get-gpo -Guid $link.gpoid
        $ou_gpo_hash.$ou_path += $gpo.DisplayName
    }
}


#report the findings

$empty_ous_without_gpos = @()

"===== ===== ===== ===== ===== ===== ===== ===== ====="
"Here are 'empty' OUs and the GPOs linked w/in them:"
foreach ($ou in ($ou_gpo_hash.Keys|sort)) {
    if ($ou_gpo_hash.Get_Item($ou)) {
        "`n"
        $ou
        foreach ($link in $ou_gpo_hash.Get_Item($ou)) {
            "..."+$link
        }
    } else {
        $empty_ous_without_gpos += $ou
    }
}

"`n"
"===== ===== ===== ===== ===== ===== ===== ===== ====="
"Here are empty OUs that do *not* contain any GPO links:"
foreach ($ou in $empty_ous_without_gpos) {
    $ou
}
