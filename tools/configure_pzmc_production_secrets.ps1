param(
    [string]$OutputPath = "$env:LOCALAPPDATA\BlankDemandPlanner\PzmcProduction\secrets.dat"
)

$ErrorActionPreference = "Stop"

function Read-SecretText([string]$Prompt) {
    $secure = Read-Host $Prompt -AsSecureString
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
}

$apiToken = Read-SecretText "API-токен ПО Производство"
$userToken = Read-SecretText "Пользовательский токен (оставьте пустым при входе по паролю)"
$password = if ([string]::IsNullOrWhiteSpace($userToken)) {
    Read-SecretText "Пароль пользователя"
}
else {
    ""
}

$payload = @{
    api_token = $apiToken
    user_token = $userToken
    password = $password
} | ConvertTo-Json -Compress

$directory = Split-Path -Parent $OutputPath
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$entropy = [Text.Encoding]::UTF8.GetBytes("BlankDemandPlanner.PzmcProduction.v1")
$clear = [Text.Encoding]::UTF8.GetBytes($payload)
$encrypted = [Security.Cryptography.ProtectedData]::Protect(
    $clear,
    $entropy,
    [Security.Cryptography.DataProtectionScope]::CurrentUser)
[IO.File]::WriteAllBytes($OutputPath, $encrypted)

Write-Host "Секреты сохранены в зашифрованном виде: $OutputPath"
