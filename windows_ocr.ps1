param(
    [Parameter(Mandatory = $true)] [string] $InputDirectory,
    [string] $Language = "ja"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
[void][Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime]
[void][Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]

$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.IsGenericMethod -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    } |
    Select-Object -First 1

function Await-WinRtOperation {
    param(
        [Parameter(Mandatory = $true)] $Operation,
        [Parameter(Mandatory = $true)] [Type] $ResultType
    )

    if ($null -eq $asTaskGeneric) {
        throw "AsTask adapter was not found"
    }
    if ($null -eq $Operation) {
        throw "WinRT operation is null"
    }
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    if ($null -eq $asTask) {
        throw "AsTask adapter could not be specialized for $ResultType"
    }
    $task = $asTask.Invoke($null, @($Operation))
    if ($null -eq $task) {
        throw "AsTask returned null for $ResultType"
    }
    $task.Wait()
    return $task.Result
}

$inputPath = (Resolve-Path -LiteralPath $InputDirectory).Path
$ocrLanguage = [Windows.Globalization.Language]::new($Language)
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($ocrLanguage)
if ($null -eq $engine) {
    throw "Windows OCR language is unavailable: $Language"
}

$results = foreach ($imageFile in Get-ChildItem -LiteralPath $inputPath -File -Filter "*.png" | Sort-Object Name) {
    $stream = $null
    $bitmap = $null
    $stage = "open storage file"
    try {
        $storageFile = Await-WinRtOperation `
            ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imageFile.FullName)) `
            ([Windows.Storage.StorageFile])
        $stage = "open image stream"
        $stream = Await-WinRtOperation `
            ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) `
            ([Windows.Storage.Streams.IRandomAccessStream])
        $stage = "create bitmap decoder"
        $decoder = Await-WinRtOperation `
            ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) `
            ([Windows.Graphics.Imaging.BitmapDecoder])
        $stage = "decode bitmap"
        $bitmap = Await-WinRtOperation `
            ($decoder.GetSoftwareBitmapAsync()) `
            ([Windows.Graphics.Imaging.SoftwareBitmap])
        $stage = "recognize bitmap"
        $ocrResult = Await-WinRtOperation `
            ($engine.RecognizeAsync($bitmap)) `
            ([Windows.Media.Ocr.OcrResult])

        [pscustomobject]@{
            file = $imageFile.Name
            text = $ocrResult.Text
            lines = @(
                foreach ($line in $ocrResult.Lines) {
                    [pscustomobject]@{
                        text = $line.Text
                        words = @(
                            foreach ($word in $line.Words) {
                                [pscustomobject]@{
                                    text = $word.Text
                                    x = [math]::Round($word.BoundingRect.X, 2)
                                    y = [math]::Round($word.BoundingRect.Y, 2)
                                    width = [math]::Round($word.BoundingRect.Width, 2)
                                    height = [math]::Round($word.BoundingRect.Height, 2)
                                }
                            }
                        )
                    }
                }
            )
            error = ""
        }
    }
    catch {
        [pscustomobject]@{
            file = $imageFile.Name
            text = ""
            lines = @()
            error = "${stage}: $($_.Exception.Message)"
        }
    }
    finally {
        if ($null -ne $bitmap) {
            $bitmap.Dispose()
        }
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

@($results) | ConvertTo-Json -Depth 8 -Compress
