
Param(
[string] $startHour,
[string] $endHour
)

$currentDate = Get-Date

$startDate = [datetime]::ParseExact($startHour, 'HH', $null)
$endDate = [datetime]::ParseExact($endHour, 'HH', $null)

if ($currentDate -gt $endDate -or $currentDate -lt $startDate)
{
   Break
}

$pythonPath="py" # todo: PYTHONPATH ?

& $pythonPath "LunchAlertor.py" $startHour $endHour
