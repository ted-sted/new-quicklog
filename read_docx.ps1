Add-Type -AssemblyName System.IO.Compression.FileSystem
$docxPath = 'c:\Users\sator\マイドライブ\Claude\QuickLog\QuickLog_仕様書_v1.0.docx'
$zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
$xmlEntry = $zip.GetEntry('word/document.xml')
$stream = $xmlEntry.Open()
$reader = New-Object IO.StreamReader($stream)
$xmlStr = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()
$xmlStr -replace '<[^>]+>', '' | Out-File -Encoding UTF8 'c:\Users\sator\マイドライブ\Claude\QuickLog\spec.txt'
