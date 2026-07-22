param(
  [Parameter(Mandatory=$true)][string[]]$Item,
  [Parameter(Mandatory=$true)][string]$Warehouse,
  [switch]$Control
)

$query = @()
foreach ($value in $Item) {
  $query += "item=$([uri]::EscapeDataString($value))"
}
$query += "warehouse=$([uri]::EscapeDataString($Warehouse))"
if ($Control) {
  $query += "control=1"
}
$url = "http://127.0.0.1:8765/stock?" + ($query -join "&")
Invoke-RestMethod -Uri $url -Method Get
