@rem -*- coding: cp932-dos -*-
Param([Parameter(ValueFromPipeline=$True)]$FilePath, $DB_FilePath, $ExportToDirectoryPath = '.\', [System.String]$Delimiter, [System.String]$Encoding = 'Default', $XlWBATemplate = -4167, [System.Int32]$XlFileFormat = 51, [switch]$SkipErrors, [switch]$ExportPlainText, [switch]$ExportCSV, [switch]$ExportJSON, [switch]$ExportClixml, [switch]$ExportXHTML, [Array]$XHTML_CSSFilePath, [Array]$XHTML_JavaScriptFilePath, [Array]$XHTML_VBScriptFilePath, [System.String]$XHTML_AnchorTarget, [switch]$ExportExcelBook, [switch]$ShowExports, [switch]$GridView, [switch]$Version)
if ($Version) {
	Write-Output 'asql Ver 1.1.26770804'
	Exit 0
}
if ([System.Environment]::Is64BitProcess) {
	$CmdLine = (($env:SystemRoot + '\syswow64\WindowsPowerShell\v1.0\powershell.exe') -Replace '\s', '` ') + ' ' + $MyInvocation.Line
	Invoke-Expression $CmdLine
	Exit $LastExitCode
}
$OutputEncoding = [System.Text.Encoding]::$Encoding
$ErrorActionPreference = "SilentlyContinue"
Set-Variable -Name APOSTROPHE -Value "'" -Option Constant
if ($FilePath.GetType().Name -Ne 'FileInfo') {
	if (Test-Path $FilePath) {
		$FilePath = Get-Item $FilePath
	} else {
		$Host.UI.WriteErrorLine("$FilePath は存在しません。")
		Exit 1
	}
}
if ($ExportToDirectoryPath.GetType().Name -Ne 'FileInfo') {
	if (Test-Path $FilePath) {
		$ExportToDirectoryPath = Get-Item $ExportToDirectoryPath
	} else {
		$ExportToDirectoryPath = Get-Item '.\'
	}
}
if (!$DB_FilePath) {
	$DB_FilePath = Join-Path $FilePath.DirectoryName ($FilePath.BaseName + '.mdb')
}
if ($DB_FilePath.GetType().Name -Ne 'FileInfo') {
	if (Test-Path -LiteralPath $DB_FilePath) {
		$DB_FilePath = Get-Item $DB_FilePath
	} else {
		$DB_FilePath = (New-Item $DB_FilePath -Type File -Force).FullName
		Remove-Item $DB_FilePath
		$DB_Catalog = New-Object -ComObject ADOX.Catalog
		$DB_Catalog.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$DB_FilePath;") | Out-Null
		$DB_Catalog = $Null
	}
}
$DB_Connection = New-Object -ComObject ADODB.Connection
$DB_Connection.Open("Provider= Microsoft.Jet.OLEDB.4.0;Data Source=$DB_FilePath")   
$DB_Command = New-Object -ComObject ADODB.Command
$DB_Command.ActiveConnection = $DB_Connection

$CommentFlag = $False
$CommandTexts = Get-Content -Path $FilePath -ReadCount 0 | foreach {
	$_ -Replace '/\*[^(?:\*/)]*?\*/|--[^(?:\*/)]*$', ''
} | foreach {
	if ($CommentFlag) {
		if ($CommentFlag = !($_ -Match '\*/')) {
			$_ -Replace '.*', ''
		} else {
			$_ -Replace '^[^(?:\*/)]*?\*/', ''
		}
	} else {
		if ($CommentFlag = $_ -Match '^[^(--)]*/\*') {
			$_ -Replace '/\*[^(?:/\*)]*$', ''
		} else {
			$_
		}
	}
} | foreach {
	$_ -Replace '--.*$', ''
} | foreach {
	$Line = $_
	while ($Line -Match "'[^';]*;[^']*'|`"[^`";]*;[^`"]*`"") {
		$Line = $Line -Replace "('[^';]*);([^']*')|(`"[^`";]*);([^`"]*`")", ('$1$3' + "`0" + '$2$4')
	}
	$_ -Replace '^.*$', $Line
}
$CommandTexts = ";`n" + ($CommandTexts -Join "`n")
while ($CommandTexts -Match ";\s*?`n\s*?`n") {
	$CommandTexts = $CommandTexts -Replace ";\s*?`n\s*?`n", ";`n;`n"
}
$CommandTexts = $CommandTexts -Split ";", 0, "simplematch"
$LineCount = 0
$Xl_Application = $Null
$Xl_TemplateSheet = $Null
$Xl_DefaultSheetsCount = $Null
$Xl_AutoFillFormulasInLists = $Null
if ($ExportExcelBook) {
	$Xl_Application = New-Object -ComObject Excel.Application
	$Xl_DisplayAlerts = $Xl_Application.DisplayAlerts
	$Xl_AutoFillFormulasInLists = $Xl_Application.AutoCorrect.AutoFillFormulasInLists
	$Xl_Application.DisplayAlerts = $False
	$Xl_Application.AutoCorrect.AutoFillFormulasInLists = $False
	$Xl_Application.ScreenUpdating = $False
	if ($XlWBATemplate.GetType().Name -Eq 'String') {
		if (Test-Path -LiteralPath $XlWBATemplate) {
			$XlWBATemplate = Get-Item $XlWBATemplate
		}
	}
	$Xl_Book = $Xl_Application.Workbooks.Add($XlWBATemplate)
	$Xl_TemplateSheet = $Xl_Book.Worksheets[1]
	$Xl_TemplateSheet.Name = 'TemplateSheet'
	$Xl_DefaultSheetsCount = $Xl_Book.Worksheets.Count
}
if ($ExportXHTML) {
	Add-Type -AssemblyName System.Web
}
function ScriptFinaly($ExitCode) {
	if ($DB_Recordset.State) {
		$DB_Recordset.Close()
		$DB_Recordset = $Null
	}
	$DB_Connection.Close()
	$DB_Connection = $Null
	if ($ExportExcelBook) {
		$ExcelQuit = $False
		if ($Xl_Book.Worksheets.Count -gt $Xl_DefaultSheetsCount) {
			$Xl_TemplateSheet.Visible = $False
			$ExportsFilePath = Join-Path $ExportToDirectoryPath "$($FilePath.BaseName)"
			$Xl_Book.SaveAs($ExportsFilePath, $XlFileFormat)
			if ($ShowExports) {
				$Xl_Application.Visible = $True
				$Xl_Book.Activate()
			} else {
				$ExcelQuit = $True
			}
		} else {
			$ExcelQuit = $True
		}
		$Xl_Application.ScreenUpdating = $True
		$Xl_Application.DisplayAlerts = $Xl_DisplayAlerts
		$Xl_Application.AutoCorrect.AutoFillFormulasInLists = $Xl_AutoFillFormulasInLists
		if ($ExcelQuit) {
			$Xl_Application.Quit()
			$Xl_Book.Close()
		}
		$Excel = $Null
	}
	Exit $ExitCode
}
foreach ($CommandText in $CommandTexts) {
	$LineCount += $CommandText.Length - $($CommandText -Replace "`n", '').Length
	if ($CommandText -Match "^\s*$") {
		continue
	}
	$DB_Command.CommandText = ($CommandText -Replace "`0", ';') + ';'
	trap [Exception] {
		$Caption = $Null
		if ($Error[0].Exception -Match "COMException: (.*)\s") {
			$Caption = $Matches[1]
		}
		$ErrorMessage = ($FilePath.Name + ':' + $LineCount + ': エラー: ' + $Caption) + "`n" + $DB_Command.CommandText + "`n^ " + $LineCount + "行目"
		$Host.UI.WriteErrorLine($ErrorMessage)
		if ($SkipErrors) {
			continue
		} else {
			ScriptFinaly(1)
		}
	}
	$DB_Recordset = $DB_Command.Execute()
	if ($DB_Recordset.State) {
		if ($Delimiter) {
			$DelimiterFlag = $False
			foreach ($Field in $DB_Recordset.Fields) {
				if ($DelimiterFlag) { Write-Host -NoNewLine $Delimiter }
				Write-Host -NoNewLine $Field.Name
				$DelimiterFlag = $True
			}
			Write-Host ''
			while (!$DB_Recordset.EOF) {
				$DelimiterFlag = $False
				foreach ($Field in $DB_Recordset.Fields) {
					if ($DelimiterFlag) { Write-Host -NoNewLine $Delimiter }
					Write-Host -NoNewLine $Field.Value
					$DelimiterFlag = $True
				}
				Write-Host ''
				$DB_Recordset.MoveNext()
			}
		} else {
			$QueryTableHeader = @()
			$QueryTableType = @()
			$QueryTableRows = New-Object PSCustomObject
			foreach ($Field in $DB_Recordset.Fields) {
				$QueryTableHeader += $Field.Name
				$QueryTableType += $Field.Type
				Add-Member -InputObject $QueryTableRows -MemberType NoteProperty -Name $Field.Name -Value $Null
			}
			$QueryTableBody = @($QueryTableRows)
			while (!$DB_Recordset.EOF) {
				$QueryTableRows = New-Object PSCustomObject
				foreach ($Field in $DB_Recordset.Fields) {
					Add-Member -InputObject $QueryTableRows -MemberType NoteProperty -Name $Field.Name -Value $Field.Value
				}
				$QueryTableBody += $QueryTableRows
				$DB_Recordset.MoveNext()
			}
			if ($QueryTableBody.Length -Gt 1) {
				$QueryTableBody[0] = $Null
			}
			Write-Output ($FilePath.Name + '.' + $LineCount)
			if ($ExportPlainText) {
				$ExportsFilePath = Join-Path $ExportToDirectoryPath "$($FilePath.Name).$LineCount.txt"
				Start-Transcript -Path $ExportsFilePath | Out-Null
				Write-Output $QueryTableBody | Format-Table -Wrap -AutoSize
				Stop-Transcript | Out-Null
				$PlainText = Get-Content -LiteralPath $ExportsFilePath | Select-Object -Index (@(19..($QueryTableBody.Length + 17)) + 17) | foreach { $_ -Replace "\0", "" }
				if (!$PlainText[0]) {
					$PlainText = Get-Content -LiteralPath $ExportsFilePath | Select-Object -Index (@(20..($QueryTableBody.Length + 18)) + 18) | foreach { $_ -Replace "\0", "" }
				}
				$PlainText | Out-File -FilePath $ExportsFilePath -Encoding $Encoding
				if ($ShowExports) {
					Start-Process $ExportsFilePath
				}
			} else {
				Write-Output $QueryTableBody | Format-Table -Wrap -AutoSize
			}
			if ($ExportCSV) {
				$ExportsFilePath = Join-Path $ExportToDirectoryPath "$($FilePath.Name).$LineCount.csv"
				$QueryTableBody | Export-Csv -Path $ExportsFilePath -Encoding $Encoding -NoTypeInformation
				if ($ShowExports) {
					Start-Process $ExportsFilePath
				}
			}
			if ($ExportJSON) {
				$ExportsFilePath = Join-Path $ExportToDirectoryPath "$($FilePath.Name).$LineCount.json"
				$QueryTableBody | ConvertTo-Json | Out-File -FilePath $ExportsFilePath -Encoding $Encoding
				if ($ShowExports) {
					Start-Process $ExportsFilePath
				}
			}
			if ($ExportClixml) {
				$ExportsFilePath = Join-Path $ExportToDirectoryPath "$($FilePath.Name).$LineCount.xml"
				$QueryTableBody | Export-Clixml -Path $ExportsFilePath -Encoding $Encoding
				if ($ShowExports) {
					Start-Process $ExportsFilePath
				}
			}
			if ($GridView) {
				$QueryTableBody | Out-GridView -Title ("$($FilePath.Name).$LineCount")
				Read-Host '続けるにはENTERキーを押して下さい'
			}
			if ($ExportXHTML) {
				$Title = "$($FilePath.Name).$LineCount"
				if ($QueryTableBody.Length -gt 1) {
					$p = '^<tr>'
					$r = '<tr>'
					$i = 1
					$QueryTableHeader | foreach {
						$p += '<td>(.*?)</td>'
						$r += '<td class="' + $QueryTableType[$i - 1] + ' ' + $QueryTableBody[1].($_).GetType().Name + '">$' + $i + '</td>'
						$i++
					}
					$p += '</tr>$'
					$r += '</tr>'
				}
				if ($XHTML_CSSFilePath) { $XHTML_CSS = "`n<style>`n" + [System.Web.HttpUtility]::HtmlEncode(($XHTML_CSSFilePath | foreach { Get-Content $_ }) -Join "`n") + "`n</style>" }
				if ($XHTML_JavaScriptFilePath) { $JScript = "`n<script type=`"text/javascript`">`n" + [System.Web.HttpUtility]::HtmlEncode(($XHTML_JavaScriptFilePath | foreach { Get-Content $_ }) -Join "`n") + "`n</script>" }
				if ($XHTML_VBScriptFilePath) { $VBScript = "`n<script type=`"text/vbscript`">`n" + [System.Web.HttpUtility]::HtmlEncode(($XHTML_VBScriptFilePath | foreach { Get-Content $_ }) -Join "`n") + "`n</script>" }
				$ExportsFilePath = Join-Path $ExportToDirectoryPath ".\$Title.xhtml"
				$("<?xml version=`"1.0`" encoding=`"$($OutputEncoding.WebName)`"?>") | Out-File -FilePath $ExportsFilePath -Encoding $Encoding
				$QueryTableBody | ConvertTo-Html -Head "<meta http-equiv=`"Content-Type`" content=`"text/html; charset=$($OutputEncoding.WebName)`" />`n<title>$Title</title>$(if ($XHTML_CSS) { $XHTML_CSS })" -PostContent "$(if ($JScript) { $JScript })$(if ($VBScript) { $VBScript })" | foreach {
					if ($p -and $r) {
						$_ -Replace $p, $r
					} else {
						$_
					}
				} | foreach {
					$_ -Replace '^<tr><th>\*</th></tr>$', ('<tr><th>' + $QueryTableHeader[0] + '</th></tr>')
				} | foreach {
					$_ -Replace '<td\s+?class="(\S+?)\sString">((https?|ftp)://[^\s]+?)</td>', ('<td class="$1 String URL"><a href="$2" target="' + $XHTML_AnchorTarget + '">$2</a></td>')
				} | Out-File -Append -FilePath $ExportsFilePath -Encoding $Encoding
				if ($ShowExports) {
					Start-Process $ExportsFilePath
				}
			}
			if ($ExportExcelBook) {
				$Xl_TemplateSheet.Copy($Xl_TemplateSheet)
				$Xl_Sheet = $Xl_Book.ActiveSheet
				$Xl_Sheet.Name = "$($FilePath.Name).$LineCount"
				$Xl_Range = $Xl_Sheet.Cells(1, 1)
				$QueryTableHeader | foreach {
					$Xl_Range.FormulaR1C1 = "$APOSTROPHE$($_)"
					$Xl_Range = $Xl_Range.Offset($Null, 1)
				}
				$Xl_Range = $Xl_Range.Offset(1, ($Xl_Range.Column - 1) * -1)
				$QueryTableBody | Where-Object { $_ } | foreach {
					$QueryTableRows = $_
					$QueryTableHeader | foreach {
						if ($QueryTableRows.($_).GetType().Name -Eq 'String') {
							$Xl_Range.FormulaR1C1 = "$APOSTROPHE$($QueryTableRows.($_))" -Replace "^$APOSTROPHE((https?|ftp)://[^\s]+)$", '=HYPERLINK("$1")'
						} else {
							$Xl_Range.FormulaR1C1 = "$($QueryTableRows.($_))"
						}
						[void]$Xl_Range.EntireColumn.AutoFit()
						$Xl_Range = $Xl_Range.Offset($Null, 1)
					}
					$Xl_Range = $Xl_Range.Offset(1, ($Xl_Range.Column - 1) * -1)
				}
			}
		}
		$DB_Recordset.Close()
		$DB_Recordset = $Null
	}
}
ScriptFinaly($LastExitCode)
