function Get-FullPath {
	param (
		[string]$filePath
	)
	if ($filePath) {
		if ($filePath -like "*%*") {
			$filePath = [Environment]::ExpandEnvironmentVariables($filePath)
		}
		if (-not [System.IO.Path]::IsPathRooted($filePath) -and $filePath -notmatch "\\|/") {
			$foundInPath = Get-Command $filePath -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source
			if ($foundInPath) { $filePath = $foundInPath }
		}
		if (Test-Path -LiteralPath $filePath) {
			return (Resolve-Path -LiteralPath $filePath).Path
		}
	}
}

function Get-FileInfo {
	param(
		[string]$filePath
	)
	$file = Get-Item $filePath -Force
	$newItem = @{}
	$newItem.Caption = $file.BaseName
	if ($file.Extension -like "*.lnk") {
		$shell = New-Object -ComObject WScript.Shell
		$shortcut = $shell.CreateShortcut($filePath)
		$newItem.Target = $shortcut.TargetPath
		if ($shortcut.Arguments) { $newItem.Arguments = $shortcut.Arguments }
		if ($shortcut.WorkingDirectory) { $newItem.WorkingDirectory = $shortcut.WorkingDirectory }
		$newItem.Icon = $shortcut.IconLocation
		if (!$shortcut.TargetPath -or ($shortcut.TargetPath -match "\{[A-F0-9\-]+\}")) {
			$shellapp = New-Object -ComObject Shell.Application
			$folder = $shellapp.Namespace((Split-Path $filePath -Parent))
			$item = $folder.ParseName((Split-Path $filePath -Leaf))
			$taget = $item.GetLink.Target
			if ($taget.IsFileSystem) {
				# Application
				$apps = $shellapp.NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items()
				$name = $item.Name
				$app = $apps | Where-Object { $_.name -like "*$name" }
				$map = @{
					"{905E63B6-C1BF-494E-B29C-65B732D3D21A}" = "%ProgramFiles%"
					"{6D809377-6AF0-444B-8957-A3773F02200E}" = "%ProgramFiles%"
					"{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}" = "%ProgramFiles(x86)%"
					"{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}" = "%ProgramData%"
				}
				$guid = [regex]::Match($app.Path, "\{[A-F0-9\-]+\}").Value
				if ($guid -and $map.ContainsKey($guid)) {
					$newItem.Target = $app.Path -replace $guid, $map[$guid]
				}
			} else {
				# System folder
				$guid = $taget.Path -replace "::", ""
				if (!(Get-FullPath (($shortcut.IconLocation -split ',')[0]))) {
					$newItem.Icon = (Get-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$guid\DefaultIcon" -ErrorAction SilentlyContinue)."(default)"
				}
				$newItem.Target = "explorer.exe"
				$newItem.Arguments = "shell:::$guid"
			}
		}

		if (!(Get-FullPath $newItem.Target)) {
			Write-Warning "$filePath - Broken shorcut"
			return
		}

		if ([int]$shortcut.WindowStyle -eq 3) {
			$newItem.WindowStyle = "Maximized"
		} elseif ([int]$shortcut.WindowStyle -eq 7) {
			$newItem.WindowStyle = "Minimized"
		}
		$bytes = [System.IO.File]::ReadAllBytes($filePath)
		if ($bytes[0x15] -band 0x20) { $newItem.RunAsAdmin = 1}
	} else {
		$newItem.Target = $file.FullName
		if ($file.VersionInfo.FileDescription) { $newItem.Caption = $file.VersionInfo.FileDescription }
	}
	return $newItem
}

function Get-DirectoryStructure {
	param (
		[string]$Path
	)
	$directory = Get-Item -Path $Path
	$result = @{
		"Caption" = $directory.Name
		"items" = @()
	}
	$files = Get-ChildItem -Path $Path -File
	foreach ($file in $files) {
		$result["items"] += (Get-FileInfo $file.FullName)
	}
	$subDirs = Get-ChildItem -Path $Path -Directory
	foreach ($subDir in $subDirs) {
		$subDirData = Get-DirectoryStructure -Path $subDir.FullName
		if ($subDirData["items"].Count -gt 0) {
			$result["items"] += $subDirData
		}
	}
	return $result
}

# -----------------------------------------------------------

$jsonFile = "qLaunch.json"
Set-Location -Path $PSScriptRoot

Write-Host "MAKE JSON FILE" -ForegroundColor Cyan
Write-Host "Choice source folder:" -ForegroundColor Green
Write-Host "1. Quick Launch"
Write-Host "2. Start Menu\Programs"
Write-Host "3. Start Menu\Programs (All Users)"
Write-Host "4. Other folder..."
Write-Host "Input number (1-4): " -ForegroundColor Green -noNewLine
switch (Read-Host) {
	'1' { $rootDir = (New-Object -ComObject Shell.Application).NameSpace("shell:Quick Launch").Self.Path }
	'2' { $rootDir = [Environment]::GetFolderPath("Programs") }
	'3' { $rootDir = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs" }
	'4' {
		$folder = (New-Object -ComObject Shell.Application).BrowseForFolder(0, "Select Folder", 0, 0)
		if (!$folder) { Exit }
		$rootDir = $folder.Self.Path
	}
	default { Exit }
}

Write-Host "Source: $rootDir"

$items = @(Get-DirectoryStructure -Path $rootDir).items
[PSCustomObject]$directoryStructure = [Ordered]@{
	"Settings" = @{ HotKeys = "Ctrl+Alt+Q" }
	"items" = $items
}

if ($items.Count -gt 0) {
	$msg = "Create new file $jsonFile"
	if (Test-Path $jsonFile) {
		Write-Host "File $jsonFile exist" -ForegroundColor Green
		Write-Host "Append new items [A] or Create New [N] JSON file (default [N])? " -ForegroundColor Green -noNewLine
		if ((Read-Host) -eq "A") {
			$msg = "Append data to file $jsonFile"
			$jsonIn = Get-Content -Path $jsonFile -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
			$directoryStructure.Settings = $jsonIn.Settings
			$directoryStructure.items += $jsonIn.items
		}

		$baseName = [System.IO.Path]::GetFileNameWithoutExtension($jsonFile)
		$ext = [System.IO.Path]::GetExtension($jsonFile)
		$bkFile = $baseName + "_" + (Get-Date -Format "yyyy.MM.dd-HH.mm.ss") + $ext
		Write-Host "Backup $jsonFile to $bkFile" -ForegroundColor Green
		Copy-Item $jsonFile $bkFile
	}

	$directoryStructure | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonFile -Encoding utf8
	Write-Host $msg -ForegroundColor Green
} else {
	Write-Warning "Valid files not found!"
}
