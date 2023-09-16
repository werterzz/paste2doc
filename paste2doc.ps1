function paste2doc {
	$currentDirectory = Get-Location
	$relativePath = $args[0]
	$fullPath = Join-Path -Path $currentDirectory -ChildPath $relativePath
	$word = New-Object -ComObject Word.Application
	$doc = $word.Documents.Add()
	$doc.Content.Paste()
	$doc.SaveAs($fullPath.toString())
	$word.Quit([ref]$false)
}

