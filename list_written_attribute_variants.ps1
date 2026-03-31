$ErrorActionPreference = 'Stop'

$writePatterns = @(
    [regex]'Properties\("([^"]+)"\)\.Value\s*(?:\+|-)?=',
    [regex]'Properties\("([^"]+)"\)\.Add\b',
    [regex]'Properties\("([^"]+)"\)\.AddRange\b',
    [regex]'Properties\("([^"]+)"\)\.Insert\b',
    [regex]'Properties\("([^"]+)"\)\.Remove\b',
    [regex]'Properties\("([^"]+)"\)\.RemoveAt\b',
    [regex]'Properties\("([^"]+)"\)\.Clear\b'
)

$setPattern = [regex]'SetADLDAPProperty\([^,]+,\s*"([^"]+)"'

$map = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]'

function Add-AttributeVariant {
    param([string]$name)
    $key = $name.ToLowerInvariant()
    if (-not $map.ContainsKey($key)) {
        $map[$key] = New-Object 'System.Collections.Generic.HashSet[string]'
    }
    $null = $map[$key].Add($name)
}

Get-ChildItem -Recurse -Filter *.vb | ForEach-Object {
    foreach ($line in Get-Content $_.FullName) {
        foreach ($pattern in $writePatterns) {
            foreach ($match in $pattern.Matches($line)) {
                Add-AttributeVariant $match.Groups[1].Value
            }
        }
        foreach ($match in $setPattern.Matches($line)) {
            Add-AttributeVariant $match.Groups[1].Value
        }
    }
}

$map.GetEnumerator() |
    Sort-Object Key |
    ForEach-Object {
        $entry = $_
        $variants = @()
        foreach ($variant in $entry.Value) {
            $variants += [string]$variant
        }
        $variants = $variants | Sort-Object
        $canonical = $variants[0]
        "{0}: {1}" -f $canonical, ([string]::Join(", ", $variants))
    }
