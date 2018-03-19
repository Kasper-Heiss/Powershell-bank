$test = Get-ADComputer pc0157076 -Properties * | select name, memberof    
$Output = foreach ($member in $test.memberof)
{
write-host $member
  New-Object -TypeName PSObject -Property @{
    Computer = $test.name
    Group = $member 
  } | Select-Object Computer,Group
}
$Output | Export-Csv .\test.csv