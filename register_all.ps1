
# For use Windows Recycle Bin
Add-Type -AssemblyName Microsoft.VisualBasic


# delete to windows Recycle Bin , 删除至Window回收站。
function Remove-Item-ToRecycleBin($Path) {
    $item = Get-Item -Path $Path -ErrorAction SilentlyContinue
    if ($item -eq $null)
    {
        Write-Error("'{0}' not found" -f $Path)
		return $false
    }
    else
    {
        $fullpath=$item.FullName
        Write-Verbose ("Moving '{0}' to the Recycle Bin" -f $fullpath)
        if ( (Test-Path -Path $fullpath -PathType Container) )
        {
           [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($fullpath,'OnlyErrorDialogs','SendToRecycleBin')
        }
        else
        {
		[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($fullpath,'OnlyErrorDialogs','SendToRecycleBin')
        }
		return $true
    }
}


#Fisrt should write paths to file .
# $pathsToWrite = @"
# D:\scs\testscs_1
# D:\scs\testscs_2
# #D:\scs\testscs_3
# D:\scs\testscs_4
# "@
$pathsToWrite = @"
D:\scs\appscs
D:\scs\dirscs
#D:\scs\ps1scs
D:\scs\filescs
"@

#$pathsFile="./shortcuts_path.txt"
$pathsFile ='./shortcuts_path_need.txt'

$scsHoldDir="D:\scs\filescs"

function WriteToFile ($filePath, $content) {
    if (!$filePath -or !$content) {
        Write-Warning "Both file path and content are required."
        return
    }

    if (!(Test-Path $filePath)) {
        Out-File -FilePath $filePath -InputObject $content
    } else {
        $choice = Read-Host "File already exists. Would you like to overwrite (O/o) or append (A/a)?"
        if ($choice.ToLower() -eq "o") {
            Out-File -FilePath $filePath -InputObject $content -Force
        } elseif ($choice.ToLower() -eq "a") {
            Add-Content -Path $filePath -Value $content
        } else {
            Write-Warning "Invalid choice. Aborting."
        }
    }
}

WriteToFile $pathsFile $pathsToWrite


#Second should take paths from file and create dirs 
# new or delete Dirs
function Op-DirsByFile($filePath, $action) {
	$newOp="new"
	##下次类似可以改进为数组，包括[Delete]or[Remove];
	$deleteOp="remove"
    # Test if the file exists
    if (!(Test-Path $filePath)) {
        Write-Warning "File not found: $filePath"
        return
    }
	
	# Test if action is valid
    if (!($action.ToLower() -in @($newOp, $deleteOp))) {
        Write-Warning "Action must be 'new' or 'delete'"
        return $false
    }
	
    # Open the file and read it line by line
    $lines = Get-Content $filePath
    foreach ($line in $lines) {
        # Ignore lines that start with '#'
        if ($line.StartsWith('#')) {
            continue
        }
        # Test if the line is a valid path
        if (!(Test-Path $line -IsValid)) {    
			Write-Warning "Invalid path: $path"
            continue
        }
        if ( $action -eq $newOp ) {
			# Test if the path exists
			if (Test-Path $line) {
				Write-Warning "Path already exists: $line"
				continue
			}
			# Create the directory
			if (New-Item -ItemType Directory -Path $line) {
				Write-Host "Directory created: $line"
			} else {
				Write-Warning "Failed to create directory: $line"
			}
		} elseif ( $action -eq $deleteOp ) {
		# Test if the path exists
			if (!(Test-Path $line)) {
				Write-Warning "Path not exists: $line"
				continue
			}
			# Delete the directory and send it to the Recycle Bin
			if ( (Remove-Item-ToRecycleBin $line)){
				Write-Output "The directory was deleted and sent to the recycle bin."
			}else{
				Write-Warning "Failed to delete directory: $line"
			}			
		} 
	}
}


#测试脚本目录生成正确与否
Op-DirsByFile $pathsFile "new"
#Op-DirsByFile $pathsFile "delete"
#Op-DirsByFile $pathsFile "remove"


#Third, should add created paths to System Path Variable.

function Op-WinPathVariableByFile {
    [CmdletBinding()]
    param (
        [string]$path,
        [string]$action
    )
    # Test if path is a valid file
    if (!(Test-Path $path -PathType Leaf)) {
        Write-Warning "Path does not exist or is not a file."
        return $false
    }
	$newOp="new"
	$deleteOp="delete"
    # Test if action is valid
    if (!($action.ToLower() -in @($newOp, $deleteOp))) {
        Write-Warning "Action must be 'new' or 'delete'"
        return $false
    }
	# Get current Path variable and display by line
	#########################
	#为了保留系统环境变量中的变量，而不是被替换成实际字符，必须使用这种方式进行获取及写入系统PATH变量。
	$pathVariable = [Environment]::GetEnvironmentVariable("Path", "Machine")
	$key = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
	$pathVariable=$key.GetValue('Path', '', 'DoNotExpandEnvironmentNames')
	#########################
	Write-Output "Previous Path variable is (split by ';',list by line): "
	$pathVariableSplit = $pathVariable.Split(';')
	$pathVariableSplit | ForEach-Object { Write-Output $_ }
	$prePathVariable=$pathVariable
    # Read file and process each line
    Get-Content $path | Foreach-Object {
        # Ignore lines starting with '#'
        if (!($_ -match '^#')) {
            # Test if path is valid
            if (!(Test-Path $_  -IsValid)) {
                Write-Warning "Path ['$_'] is not a valid path"
				continue
            }
            else {

                # Append or remove path from Path variable, depending on action
                if ($action.ToLower() -eq $newOp) {
					#dont use" $pathVariable -contains or -notcontains $_
					#ignore case $_path at begin
                    if ($pathVariableSplit -notcontains $_ ) {
						$pathVariable += ";$_"
						$pathVariableSplit = $pathVariable.Split(';')
                        Write-Output "Append : $pathVariable"
                    }
                    else {
                        Write-Warning "Path '$_' already exists in Path variable or Append ."
                    }
                }
                elseif ( $action.ToLower() -eq $deleteOp)  {
                    if ($pathVariableSplit -contains $_) {
                        #for -replace will try '\' as a regex
						$backslashFixStr=$_ -replace '\\','\\'
						#ignore case $_path at begin
						$pathVariable = $pathVariable -replace ";$($backslashFixStr)", ''
						$pathVariableSplit=$pathVariable.Split(';')
                        Write-Output "Path variable updated: $pathVariable"
                    }
                    else {
                        Write-Warning "Path '$_' does not exist in Path variable."
                    }
                }
			}
		}
	}

	# Set updated Path variable
	$realExecuteFlag=1
	if ( $realExecuteFlag -eq 1 -And $pathVariable -ne $prePathVariable ){
		[Environment]::SetEnvironmentVariable("PATH", $pathVariable, [EnvironmentVariableTarget]::Machine)
		Write-Output "`nUpdated Path variable , new Path Variable is (split by ';',list by line): "
		#$updatedPathVariable = [Environment]::GetEnvironmentVariable("Path", "Machine")
		$pathVariable_real = [Environment]::GetEnvironmentVariable("Path", "Machine")
		$key_real = Get-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
		$updatedPathVariable=$key.GetValue('Path', '', 'DoNotExpandEnvironmentNames')
		$updatedPathVariable.Split(';') | ForEach-Object { Write-Output $_ }
	}
}


Op-WinPathVariableByFile $pathsFile "new"
#Op-WinPathVariableByFile $pathsFile "delete"
#Op-WinPathVariableByFile $pathsFile "remove"

# Forth should add Shortcuts to Shortcuts Directory

function Op-ShortcutsByFile($filePath, $dirPath, $action)
{	
	$newOp="New"
	$deleteOp="Delete"
    # Test if the file exists and is a file
    if (!(Test-Path -Path $filePath -PathType Leaf))
    {
        Write-Warning "The file does not exist or is not a file: $filePath"
        return $false
    }
	# Test if action is valid
    if (!($action.ToLower() -in @($newOp, $deleteOp))) {
        Write-Warning "Action must be 'new' or 'delete'"
        return $false
    }

    # Test if the directory exists and is a directory
    if (!(Test-Path -Path $dirPath -PathType Container))
    {
        Write-Warning "The directory does not exist or is not a directory: $dirPath"
        return $false
    }

    # Read the file line by line and ignore lines starting with '#'
    $lines = Get-Content -Path $filePath | Where-Object {$_ -notmatch '^#'}

    # For each line, test if the path exists
    foreach ($line in $lines)
    {
        if (!(Test-Path -Path $line))
        {
            Write-Warning "The path does not exist: $line"
        }
        else
        {	# Generate a shortcut for the path if it exists
			$shortcutName="$($line.Split('\')[-1]).lnk"
			$shortcutPath = Join-Path -Path $dirPath -ChildPath $shortcutName
			
			if ( $action -eq $newOp ) {	
				# Test if the shortcut already exists
				if (Test-Path -Path $shortcutPath){
					Write-Warning "The shortcut already exists: $shortcutPath"
				}
				else{
					# Create the shortcut
					$WScriptShell = New-Object -ComObject WScript.Shell
					$shortcut = $WScriptShell.CreateShortcut($shortcutPath)
					$shortcut.TargetPath = $line
					$shortcut.Save()
					Write-Output "Shortcut:[shortcutPath] created."
				}
			} elseif ( $action -eq $deleteOp ) {
				# Test if the path exists
				if (!(Test-Path $shortcutPath)) {
					Write-Warning "Path not exists: $shortcutPath"
					continue
				}
				# Delete the directory and send it to the Recycle Bin
				if ( (Remove-Item-ToRecycleBin $shortcutPath)){
					Write-Output "The directory was deleted and sent to the recycle bin."
				}else{
					Write-Warning "Failed to delete directory: $shortcutPath"
				}			
			}
        }
    }
}

#测试脚本目录生成正确与否
Op-ShortcutsByFile $pathsFile $scsHoldDir "new"
#Op-ShortcutsByFile $pathsFile $scsHoldDir "delete"
