using namespace Windows.Graphics.Imaging

  # Make sure all required assemblies are loaded before any class definitions use them
  Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Runtime.WindowsRuntime
    
  # WinRT assemblies are loaded indirectly
  $null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Foundation.IAsyncOperation`1, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.RandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
    
  # Some WinRT assemblies such as [Windows.Globalization.Language] are loaded indirectly by returning the object types
  $null = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages

  #Windows.Forms
  [Windows.Forms.Application]::EnableVisualStyles()

  # Find the awaiter method
  $getAwaiterBaseMethod = [WindowsRuntimeSystemExtensions].GetMember('GetAwaiter').
  Where({$PSItem.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'}, 'First')[0]

  # Define awaiter function
  Function Await {
    param($AsyncTask, $ResultType)
    $getAwaiterBaseMethod.
        MakeGenericMethod($ResultType).
        Invoke($null, @($AsyncTask)).
        GetResult()
  }

  # Create form window
  $form = New-Object Windows.Forms.Form
  $form.Text = 'OCR Capture'
  $form.Width = 420
  $form.Height = 320
  $form.AutoSize = $true
  # Create resizeable rich textbox
  $textBox = New-Object Windows.Forms.RichTextBox
  $textBox.Multiline = $true
  $textBox.ScrollBars = [Windows.Forms.ScrollBars]::Both
  $textBox.WordWrap = $true
  $textBox.AcceptsTab = $true
  $textBox.Dock = [Windows.Forms.DockStyle]::Fill
  $textBox.Font = New-Object System.Drawing.Font('Segoe UI', 12)
  $textBox.Text = 'Scanned text will appear here.'
  $form.Controls.Add($textBox)
  # Make menu strip
  $menu = New-Object Windows.Forms.MenuStrip
  $captureButton = New-Object Windows.Forms.ToolStripMenuItem
  $captureButton.Text = 'Capture'

  $copyButton = New-Object Windows.Forms.ToolStripMenuItem
  $copyButton.Text = 'Copy'

  # Create ComboBox
  $comboBox = New-Object Windows.Forms.ToolStripComboBox
  $comboBox.Items.AddRange(@('ja-JP', 'en-US', 'zh-CH'))
  $comboBox.SelectedIndex = 0
  $comboBox.add_SelectedIndexChanged({
    # When language selection changes
    switch ($comboBox.SelectedItem) {
      'ja-JP'  { "'ja-JP'" }
      'en-US'  { "'en-US'" }
      'zh-CH'  { "'zh-CH'" }
    }
  })
  $menu.Items.Add($captureButton)
  $menu.Items.Add($copyButton)
  $menu.Items.Add($comboBox)
  $form.Controls.Add($menu)

  # Auto language detection: $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
  $languageTag = $comboBox.SelectedItem
  $language = New-Object Windows.Globalization.Language($languageTag)
  $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($language)

Function OCRCapture {
  Write-Host ('*'*40)
  # Get old clipboard
  $oldClipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
  # Reset clipboard
  [System.Windows.Forms.Clipboard]::SetText(' ')

  # Take screenshot
  if (Test-Path -Path $env:SYSTEMROOT"\System32\SnippingTool.exe") {
    # Run snipping tool (Windows 10)
    Write-Host '> Executing Snipping Tool'
    [Diagnostics.Process]::Start('SnippingTool.exe', '/clip').WaitForExit()
  } else {
    # Run snip & sketch (Windows 11)
    Write-Host '> Executing Snip & Sketch'
    Start-Process 'explorer.exe' 'ms-screenclip:' -Wait
  }
  
  # Wait for image to be copied to clipboard
  Write-Host '> Waiting for image'
  $timeout = New-TimeSpan -Seconds 10
  $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
  do {
    $clipboard = [System.Windows.Forms.Clipboard]::GetDataObject()
    Start-Sleep 0.01   # Avoid overloading the CPU
    if ($stopwatch.elapsed -gt $timeout) {
      Write-Output 'Failed to copy image to clipboard.'
      Write-Host '> Failed. Aborting...'
      [System.Windows.Forms.Clipboard]::SetDataObject($oldClipboard)
      return
    }
  } until ($clipboard.ContainsImage())
  
  # Get image
  $bmp = $clipboard.getimage()
  # Restore old clipboard
  [System.Windows.Forms.Clipboard]::SetDataObject($oldClipboard)
  
  # If softwareBitmap has a width/height under 150px, extend the image
  $minPx = 150
  if (($bmp.Height -lt $minPx) -or ($bmp.Width -lt $minPx)) {
    $nh = [math]::max($bmp.Height, $minPx)
    $nw = [math]::max($bmp.Width, $minPx)
    Write-Host ([String]::Concat('> Extending image (',$bmp.Width,',',$bmp.Height,') -> (',$nw,',',$nh,') px'))
    $graphics = [Drawing.Graphics]::FromImage(($newBmp = [Drawing.Bitmap]::new($nw, $nh)))
    $graphics.Clear($bmp.GetPixel(0, 0))
    if (($bmp.Height -lt $minPx) -and ($bmp.Width -lt $minPx)) {
      $sf = ([math]::min(([math]::floor($minPx / [math]::max($bmp.Width, $bmp.Height))), 3))
      if ($sf -gt 1) {Write-Host ([String]::Concat('> Scaling image by ',$sf,'x'))}
    } else {
      $sf = 1
    }
    $sw = ($sf * $bmp.Width)
    $sh = ($sf * $bmp.Height)
    $graphics.DrawImage($bmp, ([math]::floor(($nw-$sw)/2)), ([math]::floor(($nh-$sh)/2)), $sw, $sh)
    $bmp = $newBmp.Clone()
    $newBmp.Dispose()
    $graphics.Dispose()
  }

  # Save bmp to memory stream
  Write-Host '> Converting image format to SoftwareBitmap'
  $memStream = [IO.MemoryStream]::new()
  $bmp.Save($memStream, 'Bmp')

  # Build SoftwareBitmap
  $r = [IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($memStream)
  $params = @{
    AsyncTask  = [BitmapDecoder]::CreateAsync($r)
    ResultType = [BitmapDecoder]
  }
  $bitmapDecoder = Await @params
  $params = @{ 
    AsyncTask = $bitmapDecoder.GetSoftwareBitmapAsync()
    ResultType = [SoftwareBitmap]
  }
  $softwareBitmap = Await @params
  $memStream.Dispose()
  $r.Dispose()

  # Run OCR
  Write-Host '> Running OCR'
  (((Await $ocrEngine.RecognizeAsync($softwareBitmap)([Windows.Media.Ocr.OcrResult])).Lines |
    ForEach-Object {$_.Text}) -Join "`n")
  Write-Host '> Completed successfully'
}

  $captureButton.add_Click({
    # If result is empty, show error message
    $form.Hide()
    # Get output from OCRCapture
    $o = ((&{OCRCapture}).Trim())
    if ($o -eq '') {$textBox.Text = 'Failed to recognize text.'} else {$textBox.Text = $o}
    $form.Show()
    $form.Activate()
  })
 
  $copyButton.add_Click({
    # Copy text to clipboard
    if ($textBox.Text -ne '') {
      [System.Windows.Forms.Clipboard]::SetText($textBox.Text)
    }
  })

  # Show form
  $form.Add_Shown({$form.Activate()})
  [void] $form.ShowDialog()
