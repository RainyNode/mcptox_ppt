param(
    [string]$Path = 'F:\idmDownlowd\picture\MCP_Model_Context_Protocol_Tool_Poisoning.vsdx'
)

$ErrorActionPreference = 'Stop'
$visio = New-Object -ComObject Visio.Application
$visio.Visible = 0

try {
    $doc = $visio.Documents.Open($Path)
    $page = $doc.Pages.Item(1)
    $items = @()

    foreach ($shape in $page.Shapes) {
        $text = [string]$shape.Text
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        $cx = $shape.CellsU('PinX').ResultIU
        $cy = $shape.CellsU('PinY').ResultIU
        $w = $shape.CellsU('Width').ResultIU
        $h = $shape.CellsU('Height').ResultIU
        $items += [pscustomobject]@{
            Id = $shape.ID
            Text = ($text -replace '\s+', ' ').Trim()
            Left = $cx - ($w / 2)
            Right = $cx + ($w / 2)
            Bottom = $cy - ($h / 2)
            Top = $cy + ($h / 2)
        }
    }

    $overlaps = @()
    for ($i = 0; $i -lt $items.Count; $i++) {
        for ($j = $i + 1; $j -lt $items.Count; $j++) {
            $a = $items[$i]; $b = $items[$j]
            $xOverlap = [Math]::Min($a.Right, $b.Right) - [Math]::Max($a.Left, $b.Left)
            $yOverlap = [Math]::Min($a.Top, $b.Top) - [Math]::Max($a.Bottom, $b.Bottom)
            if ($xOverlap -gt 0.01 -and $yOverlap -gt 0.01) {
                $overlaps += [pscustomobject]@{
                    First = $a.Text
                    Second = $b.Text
                    XOverlapIn = [Math]::Round($xOverlap, 3)
                    YOverlapIn = [Math]::Round($yOverlap, 3)
                }
            }
        }
    }

    [pscustomobject]@{
        File = $Path
        TextShapeCount = $items.Count
        TextOverlapCount = $overlaps.Count
    } | Format-List
    if ($overlaps.Count -gt 0) { $overlaps | Format-Table -AutoSize }
    $doc.Close()
    if ($overlaps.Count -gt 0) { exit 2 }
}
finally {
    $visio.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($visio) | Out-Null
}
