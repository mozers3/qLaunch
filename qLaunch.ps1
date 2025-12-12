param (
	[string]$cmdLine
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Register HotKey
Add-Type -TypeDefinition @"
	using System;
	using System.Runtime.InteropServices;
	using System.Windows.Forms;

	public class HotKeyManager {
		public const int WM_HOTKEY = 0x0312;

		[DllImport("user32.dll", SetLastError=true)]
		public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

		[DllImport("user32.dll", SetLastError=true)]
		public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

		public static uint GetModifierKeys(Keys modifiers) {
			uint result = 0;
			if ((modifiers & Keys.Alt) == Keys.Alt) result |= 0x0001;
			if ((modifiers & Keys.Control) == Keys.Control) result |= 0x0002;
			if ((modifiers & Keys.Shift) == Keys.Shift) result |= 0x0004;
			return result;
		}
	}

	public class HotKeyForm : Form {
		public event EventHandler HotKeyPressed;
		protected override void WndProc(ref Message m) {
			if (m.Msg == HotKeyManager.WM_HOTKEY) {
				var handler = HotKeyPressed;
				if (handler != null) {
					handler(this, EventArgs.Empty);
				}
			}
			base.WndProc(ref m);
		}
	}
"@ -ReferencedAssemblies System.Windows.Forms

# Right mouse click for menu item
Add-Type -TypeDefinition @"
	using System.Windows.Forms;
	public class ContextMenuItem : ToolStripMenuItem {
		public MouseButtons Button { get; private set; }
		protected override void OnMouseDown(MouseEventArgs e) {
			this.Button = e.Button;
			base.OnMouseDown(e);
		}
		protected override void OnClick(System.EventArgs e) {
			base.OnClick(e);
			this.Button = MouseButtons.None;
		}
	}
"@ -ReferencedAssemblies System.Windows.Forms

# For method "ShowContextMenu"
$null = Add-Type -MemberDefinition @'
	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
'@ -Name "TrayHelper" -Namespace "Win32" -PassThru

# Extract icon from file
Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public class IconExtractor {
	[DllImport("shell32.dll", CharSet = CharSet.Auto)]
	private static extern int ExtractIconEx(string lpszFile, int nIconIndex, 
										out IntPtr phiconLarge, out IntPtr phiconSmall, uint nIcons);

	[DllImport("shell32.dll", CharSet = CharSet.Unicode)]
	private static extern int SHDefExtractIconW(string pszIconFile, int iIndex, uint uFlags,
										out IntPtr phiconLarge, out IntPtr phiconSmall, uint nIconSize);

	[DllImport("user32.dll", SetLastError = true)]
	private static extern bool DestroyIcon(IntPtr hIcon);

	public static Bitmap ExtractIconByIndex(string filePath, int iconIndex, int size = 16) {
		IntPtr hIconLarge = IntPtr.Zero;
		IntPtr hIconSmall = IntPtr.Zero;

		try {
			if (iconIndex >= 0) {
				if (ExtractIconEx(filePath, iconIndex, out hIconLarge, out hIconSmall, 1) <= 0)
					return null;
			} else {
				if (SHDefExtractIconW(filePath, iconIndex, 0, out hIconLarge, out hIconSmall, (uint)size) != 0)
					return null;
			}

			IntPtr hIcon = size <= 16 && hIconSmall != IntPtr.Zero ? hIconSmall : hIconLarge;
			if (hIcon == IntPtr.Zero)
				return null;

			using (Icon icon = Icon.FromHandle(hIcon)) {
				return new Bitmap(icon.ToBitmap(), size, size);
			}
		}
		finally {
			if (hIconLarge != IntPtr.Zero) DestroyIcon(hIconLarge);
			if (hIconSmall != IntPtr.Zero) DestroyIcon(hIconSmall);
		}
	}
}
"@ -ReferencedAssemblies System.Drawing

function Extract-Icon {
	param (
		[string]$filePath,
		[int]$iconIndex = 0,
		[int]$iconSize = 16
	)

	try {
		if ($iconIndex -eq 0) {
			$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($filePath)
			$bitmap = $icon.ToBitmap()
			if ($bitmap.Width -ne $iconSize) {
				$bitmap = New-Object System.Drawing.Bitmap($bitmap, $iconSize, $iconSize)
			}
			return $bitmap
		}
		$bitmap = [IconExtractor]::ExtractIconByIndex($filePath, $iconIndex, $iconSize)
		if ($bitmap -ne $null) {
			return $bitmap
		}
	} catch {}
}

function Get-Image {
	param (
		[string]$icon,
		[string]$tagent
	)

	if ($icon) {
		if ($icon -eq "folder") {
			return Extract-Icon -FilePath "$env:windir\system32\shell32.dll" -IconIndex 4
		} elseif ($icon -match '^(.*?)(?:,(-?\d+))?$') {
			$iconPath = Get-FullPath $matches[1]
			$iconIndex = $matches[2]
		}
		if ($iconPath -and (Test-Path -LiteralPath $iconPath)) {
			return Extract-Icon -FilePath $iconPath -IconIndex $iconIndex
		}
	}
	$tagentPath = Get-FullPath $tagent
	if ($tagentPath) {
		if ($iconIndex) {
			return Extract-Icon -FilePath $tagentPath -IconIndex $iconIndex
		} else {
			return Extract-Icon -FilePath $tagentPath
		}
	}
}

function Get-FullPath {
	param (
		[string]$filePath
	)

	if ($filePath) {
		if ($filePath -like "*%*") {
			$filePath = [Environment]::ExpandEnvironmentVariables($filePath)
		}
		try {
			if (-not [System.IO.Path]::IsPathRooted($filePath) -and $filePath -notmatch "\\|/") {
				$foundInPath = Get-Command $filePath -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source
				if ($foundInPath) { $filePath = $foundInPath }
			}
			if (Test-Path -LiteralPath $filePath) {
				return (Resolve-Path -LiteralPath $filePath).Path
			}
		} catch {
			$notifyIcon.ShowBalloonTip(2000, "Warning", $_.Exception.Message, [Windows.Forms.ToolTipIcon]::Warning)
		}
	}
}

function Run-MenuItem {
	param(
		[object]$sourceItem,
		[bool]$runAsAdmin
	)

	$filePath = Get-FullPath $sourceItem.Target
	$arguments = $sourceItem.Arguments
	$workDir = Get-FullPath $sourceItem.WorkingDirectory
	$winStyle = $sourceItem.WindowStyle
	$runAsAdmin = $runAsAdmin -xor !![int]$sourceItem.RunAsAdmin
	if (!$workDir) { $workDir = Split-Path $filePath -Parent }
	if (!$workDir) {
		$workDir = if ($PSCommandPath) {
			Split-Path $PSCommandPath
		} else {
			Split-Path ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
		}
	}
	$processParams = @{
		FilePath         = $filePath
		WorkingDirectory = $workDir
	}
	if ($arguments -and ($arguments -is [string] -or $arguments.Count -gt 0)) {
		$processParams.ArgumentList = $arguments
	}
	if ($winStyle -and ($winStyle -ne "Normal")) {
		$processParams.WindowStyle = $winStyle
	}
	if ($runAsAdmin) {
		$processParams.Verb = 'RunAs'
	}
	Start-Process @processParams
}

function Menu-MouseClick {
	param ( [object]$menuItem )

	if ($menuItem.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
		$actionMenu = New-Object System.Windows.Forms.ContextMenuStrip
		$actionMenu.ShowImageMargin = $false
		$headerLabel = New-Object System.Windows.Forms.ToolStripLabel
		$headerLabel.Text = " $($menuItem.Text)"
		$headerLabel.Image = $menuItem.Image
		$headerLabel.Font = New-Object System.Drawing.Font($headerLabel.Font, [System.Drawing.FontStyle]::Bold)
		[void]$actionMenu.Items.Add($headerLabel)
		[void]$actionMenu.Items.Add("-")
		foreach ($caption in @("Edit item …", "Delete item", "Add separator", "Add new file …")) {
			$item = New-Object System.Windows.Forms.ToolStripMenuItem
			$item.Text = "  $($caption)"
			$item.Tag = $menuItem
			[void]$actionMenu.Items.Add($item)
		}
		$actionMenu.Items[2].Add_Click{ # Edit
			Show-Form -SourceItem $this.Tag.Tag -IsEditMode $true
		}
		$actionMenu.Items[3].ForeColor = "Red"
		$actionMenu.Items[3].Add_Click{ # Delete
			Update-Collection -Collection $objJSON -FindItem $this.Tag.Tag -NewItem $null
		}
		$actionMenu.Items[4].Add_Click{ # Add separator
			$separator = [PSCustomObject]@{type = "separator"}
			Update-Collection -Collection $objJSON -FindItem $this.Tag.Tag -NewItem $separator
		}
		$actionMenu.Items[5].Add_Click{ # Add file
			$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
			$fileDialog.DereferenceLinks = $false
			$fileDialog.Title = "Select file"
			$result = $fileDialog.ShowDialog()
			if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
				$newFilePath = $fileDialog.FileName
				Show-Form -SourceItem $(Get-FileInfo -FilePath $newFilePath) -OriginalItem $this.Tag.Tag
			}
		}

		$screenPos = $menuItem.Owner.PointToScreen([System.Drawing.Point]::new($menuItem.Bounds.Right, $menuItem.Bounds.Top))
		$actionMenu.Show($screenPos)
	} else {
		$runAsAdmin = ([System.Windows.Forms.Control]::ModifierKeys -eq [System.Windows.Forms.Keys]::Shift)
		Run-MenuItem -SourceItem $menuItem.Tag -RunAsAdmin $runAsAdmin
	}
}

function Create-MenuItems {
	param (
		[System.Windows.Forms.ToolStripItemCollection]$parent,
		[object]$items
	)

	foreach ($item in $items) {
		if ($item.items) {
			$subMenu = New-Object System.Windows.Forms.ToolStripMenuItem $item.Caption
			$subMenu.Image = Get-Image -Icon "folder"
			Create-MenuItems -parent $subMenu.DropDownItems -items $item.items
			[void]$parent.Add($subMenu)
		} else {
			if ($item.type -ne "separator") {
				$target = Get-FullPath $item.Target
				if ($target) {
					$menuItem = New-Object ContextMenuItem
					$menuItem.Text = $item.Caption
					$menuItem.Image = Get-Image -Icon $item.Icon -Tagent $target
					$menuItem.Tag = $item
					$menuItem.Add_Click({Menu-MouseClick $this})
					[void]$parent.Add($menuItem)
				}
			} else {
				$menuItem = New-Object System.Windows.Forms.ToolStripSeparator
				[void]$parent.Add($menuItem)
			}
		}
	}
}

function Create-Menu {
	Copy-Item -LiteralPath $jsonFile -Destination "$jsonFile.bak"

	$menu.Items.Clear()
	[void]$menu.Items.Add("Edit JSON", $null, { Start-Process notepad.exe -Args $jsonFile })
	[void]$menu.Items.Add("-")
	Create-MenuItems -parent $menu.Items -items $objJSON.items
	[void]$menu.Items.Add("-")
	[void]$menu.Items.Add("Exit", $null, {
		[HotKeyManager]::UnregisterHotKey($hotKeyForm.Handle, $HOTKEY_ID)
		$notifyIcon.Dispose()
		[System.Windows.Forms.Application]::Exit()
	})
}

function Update-Collection {
	param(
		[object]$collection,
		[object]$findItem,
		[object]$newItem,
		[switch]$replace
	)

	function Update-CollectionItems {
		param(
			[object]$collection,
			[object]$findItem,
			[object]$newItem,
			[switch]$replace
		)

		$compare = { param($a, $b) ($a | ConvertTo-Json) -eq ($b | ConvertTo-Json) }

		if ($newItem -and !$findItem) {
			$collection.items = @($newItem) + $collection.items
			return $collection
		}

		$newCollection = [PSCustomObject]@{ items = @() }
		foreach ($item in $collection.items) {
			if ($item.items) {
				$newCollection.items += [PSCustomObject]@{
					Caption = $item.Caption
					items   = (Update-CollectionItems -Collection $item -FindItem $findItem -NewItem $newItem -Replace:$replace).items
				}
			} else {
				if ($findItem) {
					if (& $compare $findItem $item) {
						if ($newItem) {
							$newCollection.items += $newItem
							if ($replace) { continue }
						} else { continue }
					}
					$newCollection.items += $item
				}
			}
		}
		return $newCollection
	}

	$tmp = if ($objJSON.Settings) { $objJSON.Settings } else { @{HotKeys = "Ctrl+Alt+Q"} }
	$newCollection = Update-CollectionItems -Collection $collection -FindItem $findItem -NewItem $newItem -Replace:$replace
# 	$newCollection | Add-Member -NotePropertyName 'Settings' -NotePropertyValue $tmp
# 	$objJSON = $newCollection | Select-Object Settings, items
	[PSCustomObject]$objJSON = [Ordered]@{
		"Settings" = $tmp
		"items" = $newCollection.items
	}
	$objJSON | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonFile -Encoding UTF8
}

function Show-Form {
	param(
		[object]$sourceItem,
		[object]$originalItem,
		[bool]$isEditMode
	)

	function Add-Label {
		param(
			[string]$text,
			[array]$location
		)
		$label = New-Object System.Windows.Forms.Label
		$label.Text = $text
		$label.AutoSize = $true
		$label.Location = New-Object System.Drawing.Point($location[0], ($location[1]+2))
		$formEditItem.Controls.Add($label)
	}

	function Add-TextBox {
		param(
			[string]$text,
			[array]$size,
			[array]$location
		)
		$textBox = New-Object System.Windows.Forms.TextBox
		$textBox.Size = New-Object System.Drawing.Size($size[0], $size[1])
		$textBox.Location = New-Object System.Drawing.Point($location[0], $location[1])
		$formEditItem.Controls.Add($textBox)
		return $textBox
	}

	$formEditItem = New-Object System.Windows.Forms.Form
	$formEditItem.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
	$formEditItem.MaximizeBox = $false
	$formEditItem.MinimizeBox = $false
	$formEditItem.Icon = $appIcon
	$formEditItem.ClientSize = New-Object System.Drawing.Size(632, 226)

	Add-Label -Text "Caption:" -Location @(14, 18)
	Add-Label -Text "Target:" -Location @(14, 43)
	Add-Label -Text "Working Directory:" -Location @(14, 69)
	Add-Label -Text "Arguments:" -Location @(14, 95)
	Add-Label -Text "Icon:" -Location @(14, 121)
	Add-Label -Text "Window Style:" -Location @(14, 147)

	$tboxName = Add-TextBox -Size @(500) -Location @(116, 18)
	$tboxTarget = Add-TextBox -Size @(500) -Location @(116, 43)
	$tboxTarget.Add_TextChanged({ $pbIcon.Image = Get-Image -Icon $tboxIcon.Text -Tagent $tboxTarget.Text })
	$tboxWorkDir = Add-TextBox -Size @(500) -Location @(116, 69)
	$tboxArgs = Add-TextBox -Size @(500) -Location @(116, 95)
	$tboxIcon = Add-TextBox -Size @(476) -Location @(140, 121)
	$tboxIcon.Add_TextChanged({ $pbIcon.Image = Get-Image -Icon $tboxIcon.Text -Tagent $tboxTarget.Text })

	$pbIcon = New-Object System.Windows.Forms.PictureBox
	$pbIcon.Size = New-Object System.Drawing.Size(22, 22)
	$pbIcon.Location = New-Object System.Drawing.Point(114, 121)
	$pbIcon.BackColor = "ControlLightLight"
	$pbIcon.BorderStyle = "Fixed3D"
	$pbIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::CenterImage
	$formEditItem.Controls.Add($pbIcon)

	$cbWinStyle = New-Object System.Windows.Forms.ComboBox
	$cbWinStyle.Size = New-Object System.Drawing.Size(80)
	$cbWinStyle.Location = New-Object System.Drawing.Point(116, 147)
	$cbWinStyle.Items.AddRange(@("Normal", "Maximized", "Minimized", "Hidden"))
	$formEditItem.Controls.Add($cbWinStyle)

	$chkAdmin = New-Object System.Windows.Forms.CheckBox
	$chkAdmin.Text = "Run As Administrator"
	$chkAdmin.Size = New-Object System.Drawing.Size(160, 24)
	$chkAdmin.Location = New-Object System.Drawing.Point(480, 147)
	$formEditItem.Controls.Add($chkAdmin)

	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Size = New-Object System.Drawing.Size(120, 24)
	$btnCancel.Location = New-Object System.Drawing.Point(98, 186)
	$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$formEditItem.Controls.Add($btnCancel)

	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "Save"
	$btnOK.Size = New-Object System.Drawing.Size(120, 24)
	$btnOK.Location = New-Object System.Drawing.Point(414, 186)
	$btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$formEditItem.Controls.Add($btnOK)

	$tboxName.Text = $sourceItem.Caption
	$tboxTarget.Text = $sourceItem.Target
	$tboxArgs.Text = $sourceItem.Arguments
	$tboxWorkDir.Text = $sourceItem.WorkingDirectory
	$tboxIcon.Text = $sourceItem.Icon
	$cbWinStyle.SelectedIndex = if ($sourceItem.WindowStyle) {
		$cbWinStyle.FindStringExact($sourceItem.WindowStyle)
	} else { 0 }
	$chkAdmin.Checked = $sourceItem.RunAsAdmin

	$formEditItem.AcceptButton = $btnOK
	$formEditItem.CancelButton = $btnCancel

	if ($isEditMode) {
		$formEditItem.Text = "Edit Item"
		$btnCancel.Text = "Delete"
		$btnCancel.ForeColor = "Red"
		$btnCancel.Add_Click({
			Update-Collection -Collection $objJSON -FindItem $sourceItem -NewItem $null
		})
	} else {
		$formEditItem.Text = "Add to $appName"
		$btnCancel.Text = "Cancel"
	}

	$result = $formEditItem.ShowDialog()

	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		$newItem = @{
			Caption   = $tboxName.Text
			Target = $tboxTarget.Text
		}
		if ($tboxArgs.Text)                  { $newItem.Arguments = $tboxArgs.Text }
		if ($tboxWorkDir.Text)               { $newItem.WorkingDirectory = $tboxWorkDir.Text }
		if ($tboxIcon.Text)                  { $newItem.Icon = $tboxIcon.Text}
		if ($cbWinStyle.SelectedIndex -ne 0) { $newItem.WindowStyle = $cbWinStyle.Text }
		if ($chkAdmin.Checked)               { $newItem.RunAsAdmin = "1" }
		if ($isEditMode) {
			Update-Collection -Collection $objJSON -FindItem $sourceItem -NewItem $newItem -Replace
		} else {
			Update-Collection -Collection $objJSON -FindItem $originalItem -NewItem $newItem
		}
	}
}

function Get-FileInfo {
	param(
		[string]$filePath
	)

	$file = Get-Item -LiteralPath $filePath -Force
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
			[void][System.Windows.Forms.MessageBox]::Show("$filePath - Broken shorcut", $appName, "OK", "Error")
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

function Set-WindowsSettings {
	param (
		[string]$exePath
	)

	$shell = New-Object -ComObject WScript.Shell
	$sendToPath = [System.Environment]::GetFolderPath('SendTo')
	$shortcut = $shell.CreateShortcut("$sendToPath\$appName.lnk")
	$shortcut.TargetPath = $exePath
	$shortcut.IconLocation = ",0"
	$shortcut.Save()

	$startupPath = [Environment]::GetFolderPath('Startup')
	$shortcut = $shell.CreateShortcut("$startupPath\$appName.lnk")
	$shortcut.TargetPath = $exePath
	$shortcut.IconLocation = ",0"
	$shortcut.Save()

	$iconSetKey = "HKCU:\Control Panel\NotifyIconSettings" # Windows 11 only
	if (Test-Path $iconSetKey) {
		foreach ($appID in (Get-ChildItem -Path $iconSetKey -Name)) {
			$appKey = "$iconSetKey\$($appID)"
			$appPath = (Get-ItemProperty -Path $appKey -Name ExecutablePath -ErrorAction SilentlyContinue).ExecutablePath
			if ($appPath -like "*\qLaunch.exe") {
				$valueName = "IsPromoted"
				if ((Get-ItemProperty -Path $appKey -Name $valueName -ErrorAction SilentlyContinue).$valueName -eq $null) {
					Set-ItemProperty -Path $appKey -Name IsPromoted -Value 1
				}
			}
		}
	}

}

function Load-jsonFile {
	param(
		[string]$jsonFile
	)

	if (!(Test-Path -LiteralPath $jsonFile)) {
		[void][System.Windows.Forms.MessageBox]::Show("File $jsonFile not exist!", $appName, "OK", "Error")
		exit 1
	}
	$previousJson = $script:objJSON
	try {
		$jsonContent = Get-Content -Path $jsonFile -Raw -Encoding UTF8
		$script:objJSON = $jsonContent | ConvertFrom-Json -ErrorAction Stop
		return $true
	} catch {
		if ($previousJson) {
			$script:objJSON = $previousJson
			return $false
		} else {
			exit 1
		}
	}
}

function Create-NotifyIconMenu {

	function Show-ContextMenu {
		$method = $notifyIcon.GetType().GetMethod(
			"ShowContextMenu",
			[System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
		)
		$method.Invoke($notifyIcon, $null)
	}

	function Set-HotKey {
		if ($objJSON.Settings -and $objJSON.Settings.HotKeys) {
			$hotKeyForm.ShowInTaskbar = $false
			$hotKeyForm.WindowState = "Minimized"
			$hotKeyForm.Visible = $false

			$hotkeyCombination = $objJSON.Settings.HotKeys
			$modifiers = 0
			$vkCode = 0
			$hotkeyCombination.Split('+') | ForEach-Object {
				switch ($_) {
					"Ctrl"   { $modifiers += [Windows.Forms.Keys]::Control }
					"Alt"    { $modifiers += [Windows.Forms.Keys]::Alt }
					"Shift"  { $modifiers += [Windows.Forms.Keys]::Shift }
					"Win"    { $modifiers += [Windows.Forms.Keys]::LWin }
					default  { $vkCode = [Windows.Forms.Keys]::$_.Value__ }
				}
			}
			$HOTKEY_ID = 1
			[HotKeyManager]::RegisterHotKey($hotKeyForm.Handle, $HOTKEY_ID, 
				[HotKeyManager]::GetModifierKeys($modifiers), $vkCode) | Out-Null
			$hotKeyForm.Add_HotKeyPressed({ Show-ContextMenu })
		}
	}

	$menu = New-Object System.Windows.Forms.ContextMenuStrip
	Create-Menu

	$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
	$notifyIcon.Icon = $appIcon
	$notifyIcon.Text = $appName
	$notifyIcon.Visible = $true
	$notifyIcon.ContextMenuStrip = $menu # Right mouse button
	$notifyIcon.Add_MouseClick({         # Left mouse button
		param($sender, $e)
		if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Show-ContextMenu }
	})

	$hotKeyForm = New-Object HotKeyForm
	Set-HotKey

	# -- v -- JSON File Watcher -- v -- 
	$debounceTimer = New-Object System.Windows.Forms.Timer
	$debounceTimer.Interval = 500
	$debounceTimer.Enabled = $false
	$watcher = New-Object System.IO.FileSystemWatcher
	$watcher.Path = $scriptPath
	$watcher.Filter = $jsonFile
	$watcher.NotifyFilter = [System.IO.NotifyFilters]::LastWrite
	$watcher.SynchronizingObject = $menu
	$watcher.Add_Changed({
		$this.EnableRaisingEvents = $false
		$debounceTimer.Stop()
		$debounceTimer.Start()
	})
	$debounceTimer.Add_Tick({
		$debounceTimer.Stop()
		if (Load-jsonFile $jsonFile) {
			Create-Menu
			$notifyIcon.ShowBalloonTip(1000, "JSON file has been modified", "Menu content updated", [Windows.Forms.ToolTipIcon]::Info)
		} else {
			if (Test-Path -LiteralPath "$jsonFile.bak") {
				$notifyIcon.ShowBalloonTip(2000, "Invalid JSON content", "Сhanges not accepted", [Windows.Forms.ToolTipIcon]::Warning)
				Copy-Item -LiteralPath "$jsonFile.bak" -Destination $jsonFile
			} else {
				$notifyIcon.ShowBalloonTip(3000, "Invalid JSON file $jsonFile", "Please correct the content errors", [Windows.Forms.ToolTipIcon]::Error)
				Start-Process notepad.exe -Args $jsonFile
			}
		}
		$watcher.EnableRaisingEvents = $true
	})
	$watcher.EnableRaisingEvents = $true
	# -- ^ -- JSON File Watcher -- ^ -- 

	if (!$PSCommandPath) { # qLaunch.exe is running
		$timer = New-Object System.Windows.Forms.Timer
		$timer.Interval = 2000
		$timer.Add_Tick({
			$timer.Stop()
			$timer.Dispose()
			Set-WindowsSettings ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
		})
		$timer.Start()
	}

	[System.Windows.Forms.Application]::Run()
}

function Get-ScriptPath {
	if ($PSCommandPath) {
		Write-Warning "All program features are available only in the EXE version.`n`nTo compile, please run:`n  .\$(Split-Path $PSCommandPath -Leaf) COMPILE`n`n"
		return $PSScriptRoot
	} else {
		$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
		return Split-Path $exePath
	}
}

function Initialize-AppIcon {
    $iconBase64 = "AAABAAEAEBAAAAEAGABoAwAAFgAAACgAAAAQAAAAIAAAAAEAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmMzMAAAAAAAAAAAAAAABmMzMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmMzNmMzMAAAAAAABmMzNmMzMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///9mMzNmMzNmMzNmMzP///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////9mMzNmMzP///////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmMzP///////////////9mMzMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABmMzNmMzP///////9mMzNmMzMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///9mMzNmMzNmMzNmMzP///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////9mMzNmMzP///////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////////////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//6xB//+sQf//rEH736xB+Z+sQfgfrEH4H6xB+B+sQfgfrEH4H6xB+B+sQfw/rEH+f6xB//+sQf//rEH//6xB"
    $iconBytes = [Convert]::FromBase64String($iconBase64)
    $stream = [System.IO.MemoryStream]::new($iconBytes)
    return [System.Drawing.Icon]::new($stream)
}

function Compile-Script {
	if (!(Get-Module -ListAvailable -Name ps2exe)) {
		try {
			Install-Module -Name ps2exe -Scope CurrentUser -Force -ErrorAction SilentlyContinue
		} catch {
			Write-Warning $_
			return
		}
	}
	$iconPath = "$env:temp\qLaunch.ico"
	$stream = [System.IO.File]::Create($iconPath)
	$appIcon.Save($stream)
	$stream.Close()
	$version = '1.2.2'
	Invoke-PS2EXE -InputFile $PSCommandPath -x64 -noConsole -verbose -IconFile $iconPath -Title $appName -Product $appName -Copyright 'https://github.com/mozers3/qLaunch' -Company 'mozers™' -Version $version
	Remove-Item $iconPath -Force -ErrorAction SilentlyContinue
	Exit 0
}

# -----------------------------------------------------------------------------
$appName = "ps Quick Launch"
$jsonFile = "qLaunch.json"
$appIcon = Initialize-AppIcon

if ($cmdLine -eq "COMPILE") { Compile-Script }

$scriptPath = Get-ScriptPath
Set-Location -Path $scriptPath

$objJSON = @{}

if (!(Load-jsonFile $jsonFile)) {
	[void][System.Windows.Forms.MessageBox]::Show("Invalid JSON file $jsonFile", $appName, "OK", "Error")
	Start-Process notepad.exe -Args $jsonFile
	Exit 1
}

if ($cmdLine) {
	if (Test-Path -LiteralPath $cmdLine) {
		$newFile = Get-FileInfo -FilePath $cmdLine
		if ($newFile) {
			Show-Form -SourceItem $newFile
		}
	} else {
		[void][System.Windows.Forms.MessageBox]::Show("File '$cmdLine' not exist!", $appName, "OK", "Error")
	}
} else {
	Create-NotifyIconMenu
}
