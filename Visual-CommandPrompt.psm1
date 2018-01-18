<#
.SYNOPSIS
Utility Cmdlets import Visual Studio Command Prompt environmental variables

.DESCRIPTION
See:
Get-Help Import-CmdEnvironment
Get-Help Import-VisualEnvironment
#>


# .SYNOPSIS
# Import environment variables from cmd to PowerShell
# .DESCRIPTION
# Invoke the specified command (with parameters) in cmd.exe, and import any environment variable changes back to PowerShell
# .EXAMPLE
# Import-CmdEnvironment ${Env:VS90COMNTOOLS}\vsvars32.bat x86
#
# Imports the x86 Visual Studio 2008 Command Tools environment
# .EXAMPLE
# Import-CmdEnvironment ${Env:VS100COMNTOOLS}\vsvars32.bat x86_amd64
# 
# Imports the x64 Cross Tools Visual Studio 2010 Command environment
# .EXAMPLE
# Import-CmdEnvironment ${Env:VS110COMNTOOLS}\vsvars32.bat x86, x86_amd64
# 
# Imports the x64 Cross Tools Visual Studio 2012 Command environment

function Import-CmdEnvironment {
[CmdletBinding()]
param(
   [Parameter(Position=0,Mandatory=$FALSE,ValueFromPipeline=$TRUE,ValueFromPipelineByPropertyName=$TRUE)]
   [Alias("PSPath")]
   [string]$Command = "echo"
,
   [Parameter(Position=0,Mandatory=$FALSE,ValueFromRemainingArguments=$TRUE,ValueFromPipelineByPropertyName=$TRUE)]
   [string[]]$Parameters
)
   ## If it's an actual file, then we should quote it:
	if(Test-Path $Command) { $Command = "`"$(Resolve-Path $Command)`"" }
   $setRE = new-Object System.Text.RegularExpressions.Regex '^(?<var>.*?)=(?<val>.*)$', "Compiled,ExplicitCapture,MultiLine"
   $OFS = " "
   [string]$Parameters = $Parameters
   $OFS = "`n"
	## Execute the command, with parameters.
   Write-Verbose "EXECUTING: cmd.exe /c `"$Command $Parameters > nul && set`""
	## For each line of output that matches, set the local environment variable
	foreach($match in  $setRE.Matches((cmd.exe /c "$Command $Parameters > nul && set")) | Select Groups) {
      Set-Content Env:\$($match.Groups["var"]) $match.Groups["val"]
	}
}



# .SYNOPSIS
# Imports the Visual Studio Command Prompt environment
# .DESCRIPTION
# Uses the vsvars32.bat script of any valid IDE installation to import neccessary environmental variables
# .EXAMPLE
# Import-VisualEnvironment
#
# Imports the Native Tools Command environment of the most recent version found
# .EXAMPLE
# Import-VisualEnvironment -Version 12 -Architecture x64
# 
# Imports the x64 Cross Tools Command environment of Visual Studio 12 2013

function Import-VisualEnvironment()
{
[cmdletbinding()]
param(
	[Parameter(Position=0,Mandatory=$FALSE,ValueFromPipeline=$FALSE,ValueFromPipelineByPropertyName=$FALSE)]
	[ValidateSet(8,9,10,11,12,14,2005,2008,2010,2012,2013,2015)]
	[Int32]$Version = 0
	,
	[Parameter(Position=0,Mandatory=$FALSE,ValueFromPipeline=$FALSE,ValueFromPipelineByPropertyName=$FALSE)]
	[ValidateSet("X86","Amd64","Arm")]
	[System.Reflection.ProcessorArchitecture]$Host = [System.Reflection.ProcessorArchitecture]::None
	,
	[Parameter(Position=0,Mandatory=$FALSE,ValueFromPipeline=$FALSE,ValueFromPipelineByPropertyName=$FALSE)]
	[ValidateSet("X86","Amd64","Arm")]
	[System.Reflection.ProcessorArchitecture]$Target = [System.Reflection.ProcessorArchitecture]::None
)
	# Private function only used in this cmdlet
	function Get-ToolsString()
	{
		param([Int32]$Ver)
		switch ($Ver)
		{
			8 {"VS80COMNTOOLS"}
			9 {"VS90COMNTOOLS"}
			10 {"VS100COMNTOOLS"}
			11 {"VS110COMNTOOLS"}
			12 {"VS120COMNTOOLS"}
			14 {"VS140COMNTOOLS"}
			2005 {"VS80COMNTOOLS"}
			2008 {"VS90COMNTOOLS"}
			2010 {"VS100COMNTOOLS"}
			2012 {"VS110COMNTOOLS"}
			2013 {"VS120COMNTOOLS"}
			2015 {"VS140COMNTOOLS"}
			default {""}
		}
	}
	
	# Private function only used in this cmdlet
	function Get-ArchitectureString()
	{
		param([System.Reflection.ProcessorArchitecture]$HostArch
		      ,
		      [System.Reflection.ProcessorArchitecture]$TargetArch)
		switch ($Arch)
		{
			# Platform names used by MS scripts
			"x86" {"x86"}
			"x86_amd64" {"x86_amd64"}
			"x86_arm" {"x86_arm"}
			"amd64" {"amd64"}
			"amd64_arm" {"amd64_arm"}
			"amd64_x86" {"amd64_x86"}
			"arm" {"arm"}
			# Platform names added for convenience
			"Win32" {"x86"}
			"Win64" {"amd64"}
			"x64" {"amd64"}
			default {""}
		}
	}

	[String]$CommonToolsPath = ""
	if($Version)
	{
		# If the user specified a version, we should look for it
		Write-Debug "User has specified a tools version"
		
		if(($Version -le 14) -or ($Version -ge 2005 -and $Version -le 2015))
		{
			$temp = "Env:" + (Get-ToolsString -Ver $Version)
			$CommonToolsPath = (Get-ChildItem $temp).Value

			if($CommonToolsPath.Length -ne 0)
			{
				Write-Verbose ("Using tools found under " + $CommonToolsPath)
			}
			else
			{
				Write-Error "Unkown Visual Studio version"
			}
		}
		else
		{
			Add-Type -Path .\Get-VS7.cs
		}
	}
	else
	{
		# If user didn't specify a version, try finding the newest version installed
		Write-Debug "User has not specified a tools version"
		
		for($i = 14 ; $i -ge 8 ; --$i)
		{
			$temp = "Env:" + (Get-ToolsString -Ver $i)
			if(Test-Path $temp)
			{
				$CommonToolsPath = (Get-ChildItem $temp).Value
				$Version = $i
				break
			}
		}
		if($CommonToolsPath.Length -ne 0)
		{
			Write-Verbose ("Defaulting to tools found under " + $CommonToolsPath)
		}
		else
		{
			Write-Error Write-Error "No Visual Studio installation found"
		}
	}
	
	[String]$ArchitectureString = ""
	if($Architecture.Length -ne 0)
	{
		# If user specified an architecture, we should use it
		Write-Debug "User has specified an architecture"
		
		$ArchitectureString = Get-ArchitectureString -Arch $Architecture
		
		if($ArchitectureString.Length -ne 0)
		{
			Write-Verbose ("Using architecture " + $ArchitectureString)
		}
		else
		{
			Write-Error "Unkown architecture"
		}
	}
	else
	{
		# If user didn't specify an architecture, use the native tools
		Write-Debug "User has not specified an architecture"
		
		if ([Environment]::Is64BitOperatingSystem)
		{
			$ArchitectureString = "amd64"
		}
		else
		{
			$ArchitectureString = "x86"
		}
		
		Write-Verbose ("Defaulting to architecture " + $ArchitectureString)
	}
	
	# if($Version -ne 14)
	# {
	# 	Import-CmdEnvironment -Command ($CommonToolsPath + "vsvars32.bat") -Parameters $ArchitectureString
	# }
	# else
	# {
		Import-CmdEnvironment -Command ($CommonToolsPath + "..\..\VC\vcvarsall.bat") -Parameters $ArchitectureString
	# }
}
