#############################################################################################################################################
# DM_PSIncludeLib_Core.psm1
# Desktop Management PowerShell Module for Core Functions
#
# Contributing Authors: David Sitner, Steve Koehler
#
# Function Library:
#  Invoke-DMMultiThreadingEngine(): Provides a powerful and flexible looping engine to run a PowerShell scriptblock against an array of records
#  Write-DMLog()      Writes passed string to display console and to a managed log file in a way that is intended to be flexible, powerful and lightweight
#  FindDMToolsDir()   Finds folder passed as FolderName starting at the folder passed as StartFolder, and walking up the root drive
#  Get-DMScriptDir()  Returns the directory of the calling script, or %Temp% folder if called from the console
#  Get-DMScriptName() Returns the name of the calling script, or "PoshCmdConsole" if called from the console
#  Get-DMUserIDFromSID()        Returns the AD UserID given a passed string with a SID
#  Get-DMUserLogonSessionData() Returns Detailed User LogonSession Data via WMI
#  Get-DMSQLQuery()   Returns the results of a SQL query given a SQL server, SQL DB and SQL query
#  Confirm-DMCred     Returns True if the passed credential object is valid; otherwise returns False
#  Is-DMAdmin()       Returns True if the run in an administrative security context; otherwise returns False
#  Is-DMPCOnLine()    Returns True if passed PC Name is pingable and its NetBIOS name matches the passed PC Name; otherwise it returns false.
#  Parse-DMIniFile()  Imports the contents of IniFile into a hashtable
#  Split-DMArray()    Splits a passed array into the number of parts specified
#  Get-DMCName()      Converts an AD distinguished name to a Canonical Name
#  Get-DMMSIDsAssignedToEmpID() Used to get all MSIDs assigned to an EmployeeID
#  Get-DMEmpIDFromName()        Used to convert a name in various formats to an EmployeeID
#
# Change Control
#  061014 DSS V1.0 Initial Release:
#    Added: Get-DMScriptDir
#    Added: Get-DMScriptName
#    Added: Get-DMUserIDFromSID
#    Added: Write-DMLog
#  061014 DSS V1.1 Added: Find-DMToolsDir
#  071114 SJK V1.2 Fixed: Write-Progress at end of script was prompting user due to missing -Activity  switch
#  060315 DSS V1.3 Added: Get-SQLQuery
#  061115 DSS V1.4 Added: Confirm-DMCred, Renamed Get-SQLQuery to Get-DMSQLQuery, Renamed Parse-Inifile to Parse-DMInifile with Alias
#  070115 DSS V1.5 Added: Added -ErrorAction SilentlyContinue to Alias definitions
#  082515 DSS V1.6 Added: Added [AllowEmptyCollection()] to Invoke-DMMultiThreadingEngine() RecordArray param definition
#  092815 DSS V1.7 Updated functions Write-DMLog, Get-DMScriptDir and Get-DMScriptName to properly support being run from a module (or a dot-sourced library)
#  100515 DSS V1.8 Added function Run-DMElevated
#  102915 DSS V1.9 Added function Get-DMUserLogonSessionData and tweaked Get-DMSQLQuery to support calling as Run-DMSQLCmd
#  102915 DSS V1.10 Updated function Invoke-DMMultiThreadingEngine() to return a timestamp to HungJob records.
#                   Updated Get-DMScriptDir & Get-DMScriptName to persist value after running from within a script.
#  060116 DSS V1.11 Added function Split-DMArray
#  101717 DSS V1.12 Added functions Get-DMCName, Get-DMMSIDsAssignedToEmpID and Get-DMEmpIDFromName
#############################################################################################################################################

Function Split-DMArray {
<#  
  .SYNOPSIS   
    Split an array 
  .DESCRIPTION
    https://gallery.technet.microsoft.com/scriptcenter/Split-an-array-into-parts-4357dcc1
    Author Barry Chum
    Reviewed / edited by David Sitner Optum / EUTS / Plat Ops / User Mgmt Ops
  .PARAMETER inArray
   A one dimensional array you want to split
  .EXAMPLE  
   Split-array -inArray @(1,2,3,4,5,6,7,8,9,10) -parts 3
  .EXAMPLE  
   Split-array -inArray @(1,2,3,4,5,6,7,8,9,10) -size 3
#> 

  param($inArray,[int]$parts,[int]$size)
  
  if ($parts) {
    $PartSize = [Math]::Ceiling($inArray.count / $parts)
   } 
  if ($size) {
    $PartSize = $size
    $parts = [Math]::Ceiling($inArray.count / $size)
  }

  $outArray = @()
  for ($i=1; $i -le $parts; $i++) {
    $start = (($i-1)*$PartSize)
    $end = (($i)*$PartSize) - 1
    if ($end -ge $inArray.count) {$end = $inArray.count}
    $outArray+=,@($inArray[$start..$end])
   }
  return ,$outArray
 }

Function Get-DMCName { 
param([string]$DN) 
  $Parts=$DN.Split(",") 
  $NumParts=$Parts.Count 
  $FQDNPieces=($Parts -match 'DC').Count 
  $Middle=$NumParts-$FQDNPieces 
  foreach ($x in ($Middle+1)..($NumParts)) { 
    $CN+=$Parts[$x-1].SubString(3)+'.' 
    } 
  $CN=$CN.substring(0,($CN.length)-1) 
  foreach ($x in ($Middle-1)..0) {  
    $CN+="/"+$Parts[$x].SubString(3) 
   } 
Return $CN 
} 

Function Get-DMMSIDsAssignedToEmpID ($EmpID) {
#  get-aduser -Properties MemberOf,EmployeeID,manager,LastLogonDate,passwordlastset,createtimestamp,uht-IdentityManagement-AccountType -Filter {(EmployeeID -eq $EmpID) -AND (uht-IdentityManagement-AccountType -eq 'N')}
 #  also copied uht-IdentityManagement-AccountType to new attribute UHTAcctTypeID for readability
  get-aduser  -Filter {(EmployeeID -eq $EmpID)} -Properties EmployeeID,Displayname,manager,LastLogonDate,passwordlastset,createtimestamp,AccountExpirationDate,userAccountControl,uht-IdentityManagement-AccountType,uht-EmployeeStatus,uht-GLDepartmentID,uht-IdentityManagement-Mail,uht-Division,uht-InternalSegment,uht-MarketGroup |
              select *,@{name="UHTAcctTypeID";expr={$_."uht-IdentityManagement-AccountType"}},
                       @{name="EMail";expr={$_."uht-IdentityManagement-Mail"}},
                       @{name="DeptID";expr={$_."uht-GLDepartmentID"}},
                       @{name="MgrID";expr={((($_.manager).split(",")[0]).split("="))[1]}}
 }

Function Get-DMEmpIDFromName {
   <#
    .Synopsis
      Returns the EmployeeID from passed name info
    .Description
      Queries MS AD domain using supplied name info to find the EmployeeID of the matching primary account
      Requires PowerShell Active Directory module to be installed
    .Parameter MSID
      MSID of the user account
    .Parameter DisplayName
      DisplayName of the user account
    .Parameter FN
      FirstName of the user account
    .Parameter LN
      LastName of the user account
    .Parameter MI
      Middle Initial of the user account
    .EXAMPLE
      Get-DMEmpIDFromName -MSID dsitner
      Get-DMEmpIDFromName -MSID deslsadm
      Get-DMEmpIDFromName -DisplayName "sitner, david"
      Get-DMEmpIDFromName -DisplayName "sitner, david s"
      Get-DMEmpIDFromName -FN "david" -LN "sitner" -MI "S"
      Get-DMEmpIDFromName -FN "david" -LN "sitner"

      000250520
      
      All these methods will return a single EmployeeID
         
    #>
[CmdLetBinding()]
  param (
    [Parameter(Mandatory = $False, Position = 0)]
     [String]$DisplayName = $Null,
    [Parameter(Mandatory = $False)]
     [String]$MSID = $Null,
    [Parameter(Mandatory = $False)]
     [String]$FN = $Null,
    [Parameter(Mandatory = $False)]
     [String]$LN = $Null,
    [Parameter(Mandatory = $False)]
     [String]$MI = $Null)

  $arrADUserFields = "EmployeeID,uht-IdentityManagement-AccountType,DisplayName,uht-Division,uht-InternalSegment,uht-MarketGroup".Split(",")
 
  $EmpID = $Null
  If ($MSID) {
    $EmpID = (get-aduser -Filter{(Name -eq $MSID)}  -Properties EmployeeID).EmployeeID
   }

  If (!$EmpID -and $DisplayName) {
    $DisplayName = $DisplayName.Trim()
    $User = get-aduser -Filter{(DisplayName -eq $DisplayName)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
    $EmpID = $User.EmployeeID
    If (!$EmpID) {
      $DisplayNameWC = "$DisplayName*"
      $User = get-aduser -Filter{(DisplayName -like $DisplayNameWC)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
      $EmpID = $User.EmployeeID
      }
   }
  If (!$EmpID -and $LN -and $FN) {
    $DisplayName = "$LN, $FN $MI".Trim()
    $User = get-aduser -Filter{(DisplayName -eq $DisplayName)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
    $EmpID = $User.EmployeeID
    If (!$EmpID) {
      $DisplayNameWC = "$DisplayName*"
      $User = get-aduser -Filter{(DisplayName -like $DisplayNameWC)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
      $EmpID = $User.EmployeeID
      }
   }
  $EmpID = $EmpID | select -Unique | sort
  If ($EmpID.count -gt 1) {
    Write-Warning "Multiple EmployeeIDs returned from given criteria"
    $Msg = $User | select -Property $arrADUserFields | ft * | Out-String
    Write-Warning -Message $Msg

   }
  If ($EmpID.count -eq 0) {
    Write-Warning "No EmployeeIDs returned from given criteria"
   }
  Return $EmpID
 }

Function Get-DMUserInfoFromName {
   <#
    .Synopsis
      Returns the EmployeeID from passed name info
    .Description
      Queries MS AD domain using supplied name info to find the EmployeeID of the matching primary account
      Requires PowerShell Active Directory module to be installed
    .Parameter MSID
      MSID of the user account
    .Parameter DisplayName
      DisplayName of the user account
    .Parameter FN
      FirstName of the user account
    .Parameter LN
      LastName of the user account
    .Parameter MI
      Middle Initial of the user account
    .EXAMPLE
      Get-DMEmpIDFromName -MSID dsitner
      Get-DMEmpIDFromName -MSID deslsadm
      Get-DMEmpIDFromName -DisplayName "sitner, david"
      Get-DMEmpIDFromName -DisplayName "sitner, david s"
      Get-DMEmpIDFromName -FN "david" -LN "sitner" -MI "S"
      Get-DMEmpIDFromName -FN "david" -LN "sitner"

      000250520
      
      All these methods will return a single EmployeeID
    .EXAMPLE
     Get-DMUserInfoFromName -DisplayName "Sitner, David"
     Get-DMUserInfoFromName -DisplayName "Sitner, David" -ReturnMSIDs
     Get-DMUserInfoFromName -DisplayName "Sitner, David" -ReturnMSIDs -ArrayReturnAcctTypes "P","S"
     Get-DMUserInfoFromName -DisplayName "Sitner, David" -ReturnMSIDs -ArrayReturnAcctTypes "P","S" -ReturnGroupMembership
     Get-DMUserInfoFromName -DisplayName "Sitner, David" -ReturnMSIDs -ArrayReturnAcctTypes "P","S" -ReturnGroupMembership -DeNormalizeGroupMembership
     "Morin, Richard","Sitner, David" | Get-DMUserInfoFromName
     "Morin, Richard","Sitner, David" | Get-DMUserInfoFromName -ReturnMSIDs
     "Morin, Richard","Sitner, David" | Get-DMUserInfoFromName -ReturnMSIDs -ArrayReturnAcctTypes "P","S"
     "Morin, Richard","Sitner, David" | Get-DMUserInfoFromName -ReturnMSIDs -ArrayReturnAcctTypes "P","S" -ReturnGroupMembership
     "Morin, Richard","Sitner, David" | Get-DMUserInfoFromName -ReturnMSIDs -ArrayReturnAcctTypes "P","S" -ReturnGroupMembership -DeNormalizeGroupMembership
      
      All are options for this function.  To be detailed later
         
    #>
[CmdLetBinding()]
  param (
    [Parameter(ValueFromPipeline = $True, Mandatory = $False, Position = 0)]
     [String]$DisplayName = $Null,
    [Parameter(Mandatory = $False)]
     [String]$MSID = $Null,
    [Parameter(Mandatory = $False)]
     [String]$FN = $Null,
    [Parameter(Mandatory = $False)]
     [String]$LN = $Null,
    [Parameter(Mandatory = $False)]
     [String]$MI = $Null,
    [Parameter(Mandatory = $False)]
     [Switch]$ReturnMSIDs = $False,
    [Parameter(Mandatory = $False)]
     [String[]]$ArrayReturnAcctTypes = "P",
    [Parameter(Mandatory = $False)]
     [Switch]$ReturnGroupMembership = $False,
    [Parameter(Mandatory = $False)]
     [Switch]$DeNormalizeGroupMembership = $False
    )
process{      
  $arrADUserFields = "EmployeeID,uht-IdentityManagement-AccountType,DisplayName,uht-Division,uht-InternalSegment,uht-MarketGroup".Split(",")
 
  $EmpID = $Null
  If ($MSID) {
    $EmpID = (get-aduser -Filter{(Name -eq $MSID)}  -Properties EmployeeID).EmployeeID
   }

  If (!$EmpID -and $DisplayName) {
    $DisplayName = $DisplayName.Trim()
    $User = get-aduser -Filter{(DisplayName -eq $DisplayName)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
    $EmpID = $User.EmployeeID
    If (!$EmpID) {
      $DisplayNameWC = "$DisplayName*"
      $User = get-aduser -Filter{(DisplayName -like $DisplayNameWC)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
      $EmpID = $User.EmployeeID
      }
   }
  If (!$EmpID -and $LN -and $FN) {
    $DisplayName = "$LN, $FN $MI".Trim()
    $User = get-aduser -Filter{(DisplayName -eq $DisplayName)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
    $EmpID = $User.EmployeeID
    If (!$EmpID) {
      $DisplayNameWC = "$DisplayName*"
      $User = get-aduser -Filter{(DisplayName -like $DisplayNameWC)}  -Properties $arrADUserFields | where {$_."uht-IdentityManagement-AccountType" -eq "P"}
      $EmpID = $User.EmployeeID
      }
   }
  $EmpID = $EmpID | select -Unique | sort
  If ($EmpID.count -gt 1) {
    Write-Warning "Multiple EmployeeIDs returned from given criteria"
    $Msg = $User | select -Property $arrADUserFields | ft * | Out-String
    Write-Warning -Message $Msg
   }
  If ($EmpID.count -eq 0) {
    Write-Warning "No EmployeeIDs returned from given criteria"
   }
  If ($ReturnMSIDs) {
    $Results = Get-DMMSIDsAssignedToEmpID $EmpID | where {$ArrayReturnAcctTypes -contains $_.UHTAcctTypeID} | select SamAccountName, UHTAcctTypeID
    If ($ReturnGroupMembership) {
      $Results = $Results | Select @{n="Name";e={$_.SamAccountName}},UHTAcctTypeID,
                           @{n="GroupMembership";e={(get-aduser $_.SamAccountName -Properties memberof).memberof | foreach {(($_ -split ",")[0] -split "=")[1]} | sort}}
     }
    If ($DeNormalizeGroupMembership) {
    $Results = $Results |
      foreach {
        $Name = $_.Name
        $UHTAcctTypeID = $_.UHTAcctTypeID
        If ($_.GroupMembership) {
          $_.GroupMembership |
            select @{n="Name";e={$Name}},
                   @{n="UHTAcctTypeID";e={$UHTAcctTypeID}},
                   @{n="Group";e={($_)}}
         } Else {
          $hshProps = @{"Name"          = $Name
                        "UHTAcctTypeID" = $UHTAcctTypeID
                        "Group"         = $Null
                       }
          New-Object -TypeName PSObject -Property $hshProps
         }
       }
     }
   } Else {
    $Results = $EmpID
   }
  Return $Results
 }
}

Function Run-DMElevated {   
   <#
    .Synopsis
      Creates a new elevated process to run a powershell script in.
    .Description
      The newly create process elevates the script passed in -ScriptFileToElevate parameter if specified, otherwise the calling script is elevated in the new process.
      Note:
       * The script to be elevated must be on the local drive (I haven't found a trust configuration that allows network drives yet)
       * To retreive an ExitCode from the called script, the script must explicitly return a numeric exit code using a syntax like the one below:
          [Environment]::Exit(1234)
         If the script explicitly Exits with an Exitcode, the -NoExit parameter request is not honored.
    .Parameter NoExit
      Instructs the new elevated process not to automatically exit when done.
      Note:
       * If the script explicitly Exits with an Exitcode, the -NoExit parameter request is not honored.
    .Parameter DoNotWaitForElevatedProcess
      By default the function will run synchronously and wait for the process to complete before returning.
      Use this parameter to override that behavior and run asynchronously.
      Note: 
       * If the new process runs asynchronously, Exitcode of 0 will be returned by the function no matter what exitcode is returned from the new elevated process.
    .Parameter ScriptFileToElevate
      Note:
       * The script to be elevated must be on the local drive (I haven't found a trust configuration that allows network drives yet)
    .EXAMPLE
    Create a Test script to elevate:
      PS C:\> get-content "C:\Dev\ElevateTest.ps1"
        Is-DMAdmin
        Whoami
        [Environment]::Exit(1234)

    To run the test script from an un-elevated POSH session, use Run-DMElevated as shown below:
      PS C:\> $RunElevatedResult = Run-DMElevated -ScriptFileToElevate "C:\Dev\ElevateTest.ps1"

    A new elevated powershell will be created and become visible and display the lines below before exiting:
      True
      ms\DSitner

    The Exitcode will be available as the .Exitcode property of the object returned from the function
      PS C:\> $RunElevatedResult.ExitCode
      1234


    .EXAMPLE
    Create a Test script to elevate:
      PS C:\> get-content "C:\Dev\ElevateTest.ps1"
        Is-DMAdmin
        Whoami
        [Environment]::Exit(1234)

    To run the test script from an external process such as a DOS command console or SSIS job, use Run-DMElevated as shown below:

     C:\>powershell -command "& {Run-DMElevated -ScriptFileToElevate "C:\Dev\ElevateTest.ps1"}"

      Process                                 Error                                                                  ExitCode
      -------                                 -----                                                                  --------
                                              {ErrNum, ErrMsg}                                                           1234
     C:\>@echo %ErrorLevel%
      1234

    Note: 
       1) -ReturnLogfile is passed (and assumed to be $TRUE if not passed) so the calculated log file is returned
       2) $env:Temp is used as the root folder of the logfile because Write-DMLog() is called from the interactive commandline
  #>
  [CmdLetBinding()]
  param (
    [Parameter(Mandatory = $False)]
     [Switch]$NoExit = $False,
    [Parameter(Mandatory = $False)]
     [Switch]$RunInISE = $False,
    [Parameter(Mandatory = $False)]
     [Switch]$DoNotWaitForElevatedProcess = $False,
    [Parameter(Mandatory = $False)]
     [String]$ScriptFileToElevate)

  #Create $objResults to be returned
  $objReturnResult = new-object PSObject
  $objError = @{ErrNum = 0;ErrMsg = $Null}

  $arrSPArgumentList = @()
  If (($NoExit) -and (-not ($RunInISE))) {
    $arrSPArgumentList += "-NoExit"
   }

  If ($RunInISE) {
    $SPFilePath = "$psHome\powershell_ise.exe"
   } Else {
    $SPFilePath = "$psHome\powershell.exe"
   }


  #Check if we were passed an explicit script to elevate
  if ($ScriptFileToElevate) {
    #If passed script exists
    If (Test-Path $ScriptFileToElevate) {
      #Store it as the file to pass to Posh in the Start-Process Arg List
      $arrSPArgumentList += "-file `"$ScriptFileToElevate`""
     } Else {
      #Generate File not found Error
      $strErrMsg = "Warning - Unable to find supplied script file to elevate:{$ScriptFileToElevate}"
      $objError = @{ErrNum = 101;ErrMsg = $strErrMsg}
      Write-Warning $strErrMsg
     }
   } Else { #Running in the Elevate Current Script mode
    if (Get-DMScriptName -ne "PoshCmdConsole") {  
      if (-not (Is-DMAdmin)) {
        $arrSPArgumentList += = "-file `"$(Get-DMScriptDir)\$(Get-DMScriptName).ps1`""
       }
     } Else { #Running in the Elevate Current Script mode but already elevated
      $strErrMsg = "Warning - Running in the Elevate Current Script mode but process is already elevated. Not spawning new process "
      $objError = @{ErrNum = 1;ErrMsg = $strErrMsg}
      Write-Warning $strErrMsg
     }
   } 

  If ($objError.ErrNum -eq 0 ) {
    try {
      $objProcess = Start-Process -FilePath $SPFilePath -Verb RunAs -ArgumentList ($arrSPArgumentList -join " ") -ErrorAction stop -PassThru
      [Void]$objProcess.WaitForExit()
      $ExitCode = $objProcess.ExitCode
#      $ProcessResults = Invoke-DMExecutable -sExeFile $SPFilePath -sVerb RunAs -cArgs ($arrSPArgumentList -join " ")
     } Catch {
      $strErrMsg = "Error - Failed to spawn new process in Elevated Mode"
      $objError = @{"ErrNum" = 102;"ErrMsg" = $strErrMsg}
      Write-Warning $strErrMsg
     }
   }

 $objReturnResult | add-member -membertype NoteProperty -name "Process" -Value $objProcess -force
 $objReturnResult | add-member -membertype NoteProperty -name "Error" -Value $objError -force
 $objReturnResult | add-member -membertype NoteProperty -name "ExitCode" -Value $ExitCode -force
 #Return the Result
 $objReturnResult
 #Return the ExitCode
 [Environment]::Exit($ExitCode)
} 

function Invoke-DMExecutable {
    # Runs the specified executable and captures its exit code, stdout
    # and stderr.
    # Returns: custom object.
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$sExeFile,
        [Parameter(Mandatory=$false)]
        [String[]]$cArgs,
        [Parameter(Mandatory=$false)]
        [String]$sVerb
    )

    # Setting process invocation parameters.
    $oPsi = New-Object -TypeName System.Diagnostics.ProcessStartInfo
    #$oPsi.CreateNoWindow = $true
    #$oPsi.UseShellExecute = $false
    #$oPsi.CreateNoWindow = $false
    #$oPsi.RedirectStandardOutput = $true
    #$oPsi.RedirectStandardError = $true
    $oPsi.FileName = $sExeFile
    if (! [String]::IsNullOrEmpty($cArgs)) {
        $oPsi.Arguments = $cArgs
    }
    if (! [String]::IsNullOrEmpty($sVerb)) {
        $oPsi.Verb = $sVerb
    }

    # Creating process object.
    $oProcess = New-Object -TypeName System.Diagnostics.Process
    $oProcess.StartInfo = $oPsi

    # Creating string builders to store stdout and stderr.
    $oStdOutBuilder = New-Object -TypeName System.Text.StringBuilder
    $oStdErrBuilder = New-Object -TypeName System.Text.StringBuilder

    # Adding event handers for stdout and stderr.
    <#
    $sScripBlock = {
        if (! [String]::IsNullOrEmpty($EventArgs.Data)) {
            $Event.MessageData.AppendLine($EventArgs.Data)
        }
    }
    $oStdOutEvent = Register-ObjectEvent -InputObject $oProcess `
        -Action $sScripBlock -EventName 'OutputDataReceived' `
        -MessageData $oStdOutBuilder
    $oStdErrEvent = Register-ObjectEvent -InputObject $oProcess `
        -Action $sScripBlock -EventName 'ErrorDataReceived' `
        -MessageData $oStdErrBuilder
    #>

    # Starting process.
    [Void]$oProcess.Start()
#    $oProcess.BeginOutputReadLine()
#    $oProcess.BeginErrorReadLine()
    [Void]$oProcess.WaitForExit()

    # Unregistering events to retrieve process output.
#    Unregister-Event -SourceIdentifier $oStdOutEvent.Name
#    Unregister-Event -SourceIdentifier $oStdErrEvent.Name

    $oResult = New-Object -TypeName PSObject -Property ([Ordered]@{
        "ExeFile"  = $sExeFile;
        "Args"     = $cArgs -join " ";
        "ExitCode" = $oProcess.ExitCode;
        "StdOut"   = $oStdOutBuilder.ToString().Trim();
        "StdErr"   = $oStdErrBuilder.ToString().Trim()
    })

    return $oResult
}

Function Find-DMToolsDir {
   <#
    .Synopsis
      Finds folder passed as FolderName starting at the folder passed as StartFolder, and walking up the root drive.
      Returns the fullname of the path if found, null otherwise
    .Description
      Used for for finding Tools libraries in a folder structure
    .Parameter FolderName
      Folder to Find
    .Parameter StartFolder
      Folder to start from and then work up
  #>
  [CmdLetBinding()]
  param (
    [Parameter(Mandatory = $True, Position = 0)]
     [String]$FolderName,
    [Parameter(Mandatory = $True, Position = 1)]
     [String]$StartFolder)

  $Dir = Get-Item -Path $StartFolder
  Do {
    If (Test-Path "$($Dir.FullName)\$FolderName") {Return "$($Dir.FullName)\$FolderName"}
     $Dir = $Dir.parent
   } Until ([string]::IsNullOrEmpty($Dir))
 }
 
Function Get-DMScriptDir {
   <#
    .Synopsis
      Returns the directory of the calling script, or %Temp% folder if called from the console
    .Description
      If called from a script first it will cache the value for use when at the console to easily select and run only portions of a script, which would otherwise return %Temp%
      The value is stored in the global variable Global_GetDMScriptDir_ScriptDir automatically whenever the function is run from a script.
    .Parameter FlushCache
      If FlushCache is specified the global variable Global_GetDMScriptDir_ScriptDir will be deleted.
      Use this allow the function the return the value %Temp% when run from the command console
  #>
  [CmdLetBinding()]
  param (
    [Parameter(Mandatory = $False, Position = 0)]
     [switch]$FlushCache=$False)

 #If FlushCache is specified the global variable Global_GetDMScriptDir_ScriptDir will be deleted.
 If ($FlushCache) {
   Remove-Variable Global_GetDMScriptDir_ScriptDir -Scope "Global" -Force -ErrorAction SilentlyContinue
  }
     
 #Define Function name
 $FcnName = "Get-DMScriptDir"
 #Define default value to start with
 $ScriptDir=$env:Temp
 
 #Change ErrorAction to SilentlyContinue to accomadate POSH v2.0 and avoid unneeded error messages
 $EATemp = $ErrorActionPreference
 $ErrorActionPreference = "SilentlyContinue"

 #Analyze system state and calculate needed Vars
 $MyInvocationPath = $script:MyInvocation.MyCommand.Path
 $arrCallStack = Get-PSCallStack
 $FirstNonNullScriptNameInCallStack = (($arrCallStack | Where {$_.Scriptname})[-1]).scriptname
 #$NonCmdConsoleScriptName = [io.path]::GetFileNameWithoutExtension($FirstNonNullScriptNameInCallStack)
 $NonCmdConsoleScriptDir = Split-path $FirstNonNullScriptNameInCallStack -Parent

 #Restore original ErrorAction
 $ErrorActionPreference = $EATemp

 #Write-Host '1***Get-PSCallStack  | ft Command,Scriptname -AutoSize'
 #$arrCallStack = Get-PSCallStack 
 #Get-PSCallStack | ft Command,Scriptname -AutoSize
 #Write-Host "2***MyInvocationPath: {$MyInvocationPath}"
 
 #If function is being called via a module...
 If ((get-command $FcnName).modulename) {
   #Below is when called from console and function is in a module
   #  Get-PSCallStack  | ft Command,Scriptname -AutoSize 
   #  Command         ScriptName                                                       
   # -------         ----------                                                       
   # Get-DMScriptDir C:\Windows\DES_Tools\PSModules\_DMPSLib\DM_PSIncludeLib_Core.psm1
   # <ScriptBlock>                                                                    

   # If FirstNonNullScriptNameInCallStack exists and $NonCmdConsoleScriptDir isn't found in the PS module path
   If (($FirstNonNullScriptNameInCallStack) -AND (($env:PSModulePath -split ";" | where {$NonCmdConsoleScriptDir -like "$_*"}) -eq $Null)) {
     $ScriptDir = $NonCmdConsoleScriptDir
    }
  } Else { #Else function is NOT being called via a module
  #if $MyInvocationPath exists (it will be empty if function is loaded via library or script and then function is called at the Cmd console `
    # and $NonCmdConsoleScriptDir isn't found in the PS module path
   if (($MyInvocationPath) -and (($env:PSModulePath -split ";" | where {$NonCmdConsoleScriptDir -like "$_*"}) -eq $Null)) {
     $ScriptDir = $NonCmdConsoleScriptDir
    }
  }
 #If Not Running in console
 If ($ScriptDir -ne $env:Temp) {
   #If the Global Var Already Exists, update it
   If ($Global_GetDMScriptDir_ScriptDir) {
     $Global_GetDMScriptDir_ScriptDir = $ScriptDir
    } Else { #Create and set it
     New-Variable -Name Global_GetDMScriptDir_ScriptDir -Value $ScriptDir -Scope "Global" -Option ReadOnly
    }
  } Else { #Running in Cmd Console
   #If the Global Var Already Exists, update it
   If ($Global_GetDMScriptDir_ScriptDir) {
     $ScriptDir = $Global_GetDMScriptDir_ScriptDir
    }
  }
Return $ScriptDir
 }

Function Get-DMScriptName {
   <#
    .Synopsis
      Returns the name of the calling script, or "PoshCmdConsole" if called from the console
    .Description
      If called from a script first it will cache the value for use when at the console to easily select and run only portions of a script, which would otherwise return "PoshCmdConsole"
      The value is stored in the global variable Global_GetDMScriptName_ScriptName automatically whenever the function is run from a script.
    .Parameter FlushCache
      If FlushCache is specified the global variable Global_GetDMScriptName_ScriptName will be deleted.
      Use this allow the function the return the value "PoshCmdConsole" when run from the command console
  #>
  [CmdLetBinding()]
  param (
    [Parameter(Mandatory = $False, Position = 0)]
     [switch]$FlushCache=$False)

 #If FlushCache is specified the global variable Global_GetDMScriptName_ScriptName will be deleted.
 If ($FlushCache) {
   Remove-Variable Global_GetDMScriptName_ScriptName -Scope "Global" -Force -ErrorAction SilentlyContinue
  }

 #Define Function name
 $FcnName = "Get-DMScriptName"
 #Define default value to start with
 $PoshCmdConsole = "PoshCmdConsole"
 $ScriptName = $PoshCmdConsole
 
 #Change ErrorAction to SilentlyContinue to accomadate POSH v2.0 and avoid unneeded error messages
 $EATemp = $ErrorActionPreference
 $ErrorActionPreference = "SilentlyContinue"

 #Analyze system state and calculate needed Vars
 $MyInvocationPath = $script:MyInvocation.MyCommand.Path
 $arrCallStack = Get-PSCallStack
 $FirstNonNullScriptNameInCallStack = (($arrCallStack | Where {$_.Scriptname})[-1]).scriptname
 $NonCmdConsoleScriptName = [io.path]::GetFileNameWithoutExtension($FirstNonNullScriptNameInCallStack)
 $NonCmdConsoleScriptDir = Split-path $FirstNonNullScriptNameInCallStack -Parent

 #Restore original ErrorAction
 $ErrorActionPreference = $EATemp
  
 #If function is being called via a module...
 If ((get-command $FcnName).modulename) {
   #Below is when called from console and function is in a module
   #  Get-PSCallStack  | ft Command,Scriptname -AutoSize 
   #  Command         ScriptName                                                       
   # -------         ----------                                                       
   # Get-DMScriptDir C:\Windows\DES_Tools\PSModules\_DMPSLib\DM_PSIncludeLib_Core.psm1
   # <ScriptBlock>                                                                    

   # If FirstNonNullScriptNameInCallStack exists and $NonCmdConsoleScriptDir isn't found in the PS module path
   If (($FirstNonNullScriptNameInCallStack) -AND (($env:PSModulePath -split ";" | where {$NonCmdConsoleScriptDir -like "$_*"}) -eq $Null)) {
     $ScriptName = $NonCmdConsoleScriptName
    }
  } Else { #Else function is NOT being called via a module
  #if $MyInvocationPath exists (it will be empty if function is loaded via library or script and then function is called at the Cmd console `
    # and $NonCmdConsoleScriptDir isn't found in the PS module path
   if (($MyInvocationPath) -and (($env:PSModulePath -split ";" | where {$NonCmdConsoleScriptDir -like "$_*"}) -eq $Null)) {
     $ScriptName = $NonCmdConsoleScriptName
    }
  }
 #If Not Running in console
 If ($ScriptName -ne $PoshCmdConsole) {
   #If the Global Var Already Exists, update it
   If ($Global_GetDMScriptName_ScriptName) {
     $Global_GetDMScriptName_ScriptName = $ScriptName
    } Else { #Create and set it
     New-Variable -Name Global_GetDMScriptName_ScriptName -Value $ScriptName -Scope "Global" -Option ReadOnly
    }
  } Else { #Running in Cmd Console
   #If the Global Var Already Exists, update it
   If ($Global_GetDMScriptName_ScriptName) {
     $ScriptName = $Global_GetDMScriptName_ScriptName
    }
  }
 Return $ScriptName
}

Function Get-DMUserIDFromSID {
   <#
    .Synopsis
      Returns the AD UserID given a passed string with a SID
    .Description
      Uses Get-ADUser and requires the Active-Directory PowerShell module to be installed
    .Parameter SID
      SID account
  #>
 param([string]$SID)
 Return (Get-aduser -Filter {sid -eq $sid}).Name
}

Function Get-DMUserLogonSessionData {
<#
    .Synopsis
      Returns Detailed User LogonSession Data via WMI
    .Description
      Joins the WMI classes Win32_LoggedOnUser and Win32_LogonSession on LogonID to return detailed User LogonSession Data
       on the local host or on a remote host if the -ComputerName parameter is specified
      See http://msdn.microsoft.com/en-us/library/aa394189.aspx for details on Windows LogonTypes
    .Parameter ComputerName
      Optional name of the computer to run against.
      Local host is the default
    .Example
      Get-DMUserLogonSessionData | format-table * -autosize

      ComputerName    LogonUser                       LogonTypeCode LogonTypeName AuthenticationPackage StartTime             LogonID   
      ------------    ---------                       ------------- ------------- --------------------- ---------             -------   
      LH7U05CB2221YP8 LH7U05CB2221YP8\SYSTEM                      0 System        Negotiate             10/13/2015 9:06:47 AM 999       
      LH7U05CB2221YP8 LH7U05CB2221YP8\NETWORK SERVICE             5 Service       Negotiate             10/13/2015 9:06:52 AM 996       
      LH7U05CB2221YP8 LH7U05CB2221YP8\LOCAL SERVICE               5 Service       Negotiate             10/13/2015 9:06:53 AM 997       
      LH7U05CB2221YP8 LH7U05CB2221YP8\ANONYMOUS LOGON             3 Network       NTLM                  10/13/2015 9:07:51 AM 550062    
      LH7U05CB2221YP8 MS\dsitner                                  2 Interactive   Kerberos              10/13/2015 9:08:38 AM 933565    
      LH7U05CB2221YP8 LH7U05CB2221YP8\SYSTEM                      3 Network       Kerberos              10/20/2015 1:19:55 AM 1513017174
      LH7U05CB2221YP8 MS\dsitner                                  2 Interactive   Kerberos              10/28/2015 7:43:03 PM 3942736838
      LH7U05CB2221YP8 MS\dsitner                                  3 Network       Kerberos              10/29/2015 1:18:15 PM 4068558620

    .Example
      Get-DMUserLogonSessionData -ComputerName ve8s00000080 | format-table * -AutoSize

      ComputerName LogonUser                    LogonTypeCode LogonTypeName AuthenticationPackage StartTime              LogonID 
      ------------ ---------                    ------------- ------------- --------------------- ---------              ------- 
      VE8S00000080 VE8S00000080\SYSTEM                      0 System        Negotiate             10/29/2015 12:36:19 AM 999     
      VE8S00000080 VE8S00000080\NETWORK SERVICE             5 Service       Negotiate             10/29/2015 12:36:24 AM 996     
      VE8S00000080 VE8S00000080\DWM-1                       2 Interactive   Negotiate             10/29/2015 12:36:24 AM 72107   
      VE8S00000080 VE8S00000080\DWM-1                       2 Interactive   Negotiate             10/29/2015 12:36:24 AM 72125   
      VE8S00000080 VE8S00000080\LOCAL SERVICE               5 Service       Negotiate             10/29/2015 12:36:24 AM 997     
      VE8S00000080 VE8S00000080\ANONYMOUS LOGON             3 Network       NTLM                  10/29/2015 12:36:48 AM 214241  
      VE8S00000080 MS\dsitner                               2 Interactive   Kerberos              10/29/2015 11:18:02 AM 16805737
      VE8S00000080 MS\dsitner                               2 Interactive   Negotiate             10/29/2015 11:18:02 AM 16806462
      VE8S00000080 MS\deslsadm                              2 Interactive   Kerberos              10/29/2015 2:17:23 PM  21939436
      VE8S00000080 MS\deslsadm                              2 Interactive   Negotiate             10/29/2015 2:17:23 PM  21939468
      VE8S00000080 MS\dsitner                               3 Network       Kerberos              10/29/2015 4:40:16 PM  26381478
      VE8S00000080 MS\dsitner                               3 Network       Kerberos              10/29/2015 4:40:17 PM  26381779
      VE8S00000080 MS\dsitner                               3 Network       NTLM                  10/29/2015 4:56:42 PM  26867787

  #>
  Param (
	    [Parameter(Mandatory = $False, Position = 0, ValueFromPipeline = $true)]
	     [String]$ComputerName = "$env:computerName"
	   )
  #Create hash table of LogonTypeCode to LogonTypeName mappings
  $hshLogonType = @{0="System"
                  1="N/A"
                  2="Interactive"
                  3="Network"
                  4="Batch"
                  5="Service"
                  6="Proxy"
                  7="Unlock"
                  8="NetworkCleartext"
                  9="NewCredentials"
                  10="RemoteInteractive"
                  11="CachedInteractive"
                  12="CachedRemoteInteractive"
                  13="CachedUnlock"
                 }
  $arrWMILoU = get-wmiobject -class Win32_LoggedOnUser -ComputerName $ComputerName
  $arrWMILoS = get-wmiobject -class Win32_LogonSession -ComputerName $ComputerName
  $arrLoggedOnUserSessions = $arrWMILoU |
    select @{name="ComputerName";expression={$_.__Server}}, `
    @{name="LogonUser";expression={$arrTemp=($_.Antecedent -split '"');"$($ArrTemp[1])\$($ArrTemp[3])"}}, `
    @{name="LogonTypeCode";expression={$LoUID=($_.Dependent -split '"')[1];($arrWMILoS|?{$_.LogonID -eq $LoUID}).LogonType}}, `
    @{name="LogonTypeName";expression={$LoUID=($_.Dependent -split '"')[1];$hshLogonType.[int](($arrWMILoS|?{$_.LogonID -eq $LoUID}).LogonType)}}, `
    @{name="AuthenticationPackage";expression={$LoUID=($_.Dependent -split '"')[1];($arrWMILoS|?{$_.LogonID -eq $LoUID}).AuthenticationPackage}}, `
    @{name="StartTime";expression={$LoUID=($_.Dependent -split '"')[1];[System.Management.ManagementDateTimeconverter]::ToDateTime(($arrWMILoS|?{$_.LogonID -eq $LoUID}).StartTime)}}, `
    @{name="LogonID";expression={($_.Dependent -split '"')[1]}}

  Return $arrLoggedOnUserSessions | sort StartTime
 }

Function Get-DMSQLQuery {
     <#
    .Synopsis
      Returns the results of the passed SQL Query on the passed SQL server against the passed SQL DB
       or
      Runs the passed SQL Cmd on the passed SQL server against the passed SQL DB
    .Description
      Used to run a SQL Cmd or SQL Query against a SQL DB using Pass-through Integrated Security
      Use of alias Run-DMSQLCmd is recommended when the query is really a cmd to improve code readability
      You can use the -SQLQuery parameter or the -SQLCmd parameter but not both - again for readability.
    .Parameter SQLServer
      Name of the SQL Server
    .Parameter SQLDBName
      Name of the SQL database on the SQL query
    .Parameter SQLQuery
      SQL query to run against the DB
    .Parameter SQLCmd
      SQL cmd to run against the DB
    .EXAMPLE
      Get-DMSQLQuery -SQLServer DMI-SQL-Prod -SQLDBName UHTLogonScripts -SQLQuery "SELECT top 3 * FROM [dbo].[tblLogonData_Hist] where OS not like '%SVR%' Order by Logon_ID desc"
      Run-DMSQLCmd   -SQLServer DMI-SQL-Prod -SQLDBName UHTLogonScripts -SQLCmd   "SELECT top 3 * FROM [dbo].[tblLogonData_Hist] where OS not like '%SVR%' Order by Logon_ID desc"
      Get-DMSQLQuery "DMI-SQL-Prod"  "UHTLogonScripts" "SELECT top 3 * FROM [dbo].[tblLogonData_Hist] where OS not like '%SVR%' Order by Logon_ID desc"

      These 3 commands represent 3 different syntax to query for the top 3 recent user logons
    #>
[CmdLetBinding()]
  param (
    [Parameter(Mandatory = $True, Position = 0)]
     [String]$SQLServer,
    [Parameter(Mandatory = $True, Position = 1)]
     [String]$SQLDBName,
    [Parameter(Mandatory = $True, Position = 2,ParameterSetName="Query")]
     [String]$SQLQuery,
    [Parameter(Mandatory = $False, Position = 3)]
     [int]$Timeout,
    [Parameter(Mandatory = $True,ParameterSetName="Cmd")]
     [String]$SQLCmd)

  $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
  $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True" 
  $SqlCmdObj = New-Object System.Data.SqlClient.SqlCommand
  $SqlCmdObj.CommandText = If ($SQLQuery) {$SQLQuery} Else {$SQLCmd}
  $SqlCmdObj.Connection = $SQLConnection
  If ($Timeout) {$SqlCmdObj.CommandTimeout = $Timeout}
  $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
  $SqlAdapter.SelectCommand = $SqlCmdObj 
  $DataSet = New-Object System.Data.DataSet
  $x = $SqlAdapter.Fill($DataSet) 
  $SqlConnection.Close()
  
  Return $DataSet.Tables[0]
 }

Function Confirm-DMCred {
<#
.SYNOPSIS
  Returns True if the passed credential object is valid; otherwise returns False
.DESCRIPTION
  If a script uses Get-Credential and prompts the user for password, Get-Credential will return a PSCredential object
   regardless of whether the ID was valid or the password was correct for the ID.
  This function invokes the Windows OS utility whoami.exe to confirm that a process can be launched 
   and that it runs under the security context of the passed credential object.
.EXAMPLE
  $MyCred = Get-Credential
  If (Confirm-DMCred -Cred $MyCred) {"Your Credential for $(($MyCred).username) is Valid"} Else {"Your Credential for $(($MyCred).username) is Invalid"}

  Run the code above to test the validity of the credential you are prompted for
#>
  param (
    [Parameter(Mandatory = $True, Position = 0)]
      [PSCredential]$Cred
      )
   return [boolean]((Invoke-Command -ScriptBlock{whoami} -ComputerName localhost -Credential $Cred -ErrorAction SilentlyContinue) -eq $Cred.UserName)
 }

Function Is-DMAdmin {
<#
.SYNOPSIS
  Returns True if the run in an administrative security context; otherwise returns False.
  If UAC is enabled, that means the account is an admin on the computer AND the process permissions have been elevated.
.DESCRIPTION
  v1.0 Taken from the net:
  #http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/11/check-for-admin-credentials-in-a-powershell-script.aspx
.EXAMPLE
  If (Is-DMAdmin) {"You're currently running as an Admin"} Else {"You're currently not running as an Admin"}
#>
 ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

Function Is-DMPCOnLine {
<#
.SYNOPSIS
  Returns True if passed PC Name is pingable and it's NetBIOS name matches the passed PC Name; otherwise it returns false.
.DESCRIPTION
  This function is more robust than a simple ping check
   in that it uses NBTStat command to additionally confirm the host responding the ping has the same NetBIOS name as the passed PC name.
.EXAMPLE
  If (Is-DMPCOnline "LH7U05CB2221YP8") {"PC Is online"} Else {"PC is offline"}
#>
  [CmdLetBinding()]
  param (
    [Parameter(Mandatory = $True, Position = 0)]
     [String]$PC
     )
  
  $Result = $False
  $IP = (Test-Connection -ComputerName $PC -Count 1 -ErrorAction SilentlyContinue).IPV4Address
  If ($IP -IS "IPAddress") {
    $arrNBTMatches = nbtstat -A $IP | where {$_ -like  "*$PC*<00>*"}
    If ($arrNBTMatches) {$Result = $True}
   }
  Return $Result
 }

Function Parse-DMIniFile {
<#
.SYNOPSIS
  Imports the contents of IniFile into a hashtable
.DESCRIPTION
  v1.1 Updated to allow Values to contain "=" char
  v1.0 Taken from the net:
  http://blogs.technet.com/b/heyscriptingguy/archive/2011/08/20/use-powershell-to-work-with-any-ini-file.aspx
.PARAMETER IniFile
  Path of the ini file to import
.NOTES
  Function ouputs a hash table with all ini contents
.EXAMPLE
  Parse-DMIniFile .\ScriptOptions.ini
#>
 [CmdletBinding()]
 Param (
   [Parameter(mandatory=$true,ValueFromPipeline=$true)]
   [string]$IniFile
  )
 Process {
   $ini = @{}
   switch -regex -file $IniFile {
     "^\[(.+)\]$"{
       $section = $matches[1]
       $ini[$section] = @{}
      }
     "(.+?)=(.+)" {
       $name,$value = $matches[1..2]
       $ini[$section][$name] = $value
      }
    }
   return $ini
  }
} 

Function Write-DMLog {
   <#
    .SYNOPSIS
      Writes passed string to display console and to a managed log file in a way that is intended to be flexible, powerful and lightweight.
    .DESCRIPTION
      PowerShell script logging function with usable default options for most cases.  Provides flexibility to modifying defaults through optional arguments.
      Note: 
       Debugging of Write-DMLog() can be enabled by setting $DM_DbgLog=$True
    .PARAMETER Text
      Text string to write to console display and to the log file.
    .PARAMETER LogFile
      Path to the log file
      Options: If passed LogFile is a valid absolute path, it is used as passed.
        If not, it is interpreted as a -childpath to the Join-Path commandlet, using a root folder of <ScriptDir>\Logs - where <ScriptDir> is the folder containing the running script.
        If the Write-DMLog is called from the interactive commandline, <ScriptDir> is assigned the $env:Temp folder (by DM Library function Get-DMScriptDir())
    .PARAMETER CacheLogFileName
      CacheLogFileName is an integer which is used to represent 3 states: True, False and UnDefined(default).
      It does this by being set to 1 of 3 values: 1=True, 0=False and -1=Undefined
       If set to True via Syntax: -CacheLogFileName $True
          the LogFile name is cached and and saved in the the global variable: $Global_WriteDMLog_LogFile
       If set to Undefined(default value if parameter is not specified)
          the cached LogFile name is used if $Global_WriteDMLog_LogFile has been set previously, otherwise the logfile is calculated 
      Note: 
       This is implemented as a TriState integer so that True/False logic can be used when passing and scripting with the parameter.
       This technique works because [boolean]1 = [boolean]-1 = True, and [boolean]0 = False
       So CacheLogFileName
    .PARAMETER UseCacheLogFileName
      To increase performance the LogFile name is cached and not recalculated each time the function is called.
      However if within a script multiple log files are to be written to, use this switch -CacheLogFileName:$false  
    .PARAMETER UseDateBasedLogFiles
      Overrides the default naming convention for Logfiles and uses a DateBased LogFile format of:
      <ScriptDir>\Logs\<ScriptName>-yyyyMMddhhmmtt
      For Example: C:\ScriptDir\Logs\ScriptName-201410020639PM.log
    .PARAMETER DateBasedLogFileFormat
      Overrides the default DateBased LogFile format of <ScriptDir>\Logs\<ScriptName>-yyyyMMddhhmmtt
      For options see: http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo(VS.85).aspx
    .PARAMETER DateBasedLogFileAgeInDays
      The age threshold (in days) of the oldest log file to keep.
      This parameter is only considered if called at the same time as the ManageLogFiles is passed as True.
      If ManagerLogFiles is Ture, older log files will be deleted.
    .PARAMETER ManageLogFiles
      If True, log files will be pruned to keep file size below MaxLogFileSize (argument), previous log files will be renamed by appending an _# to the name, and deleted to keep the log files under MaxLogFilesToKeep (argument).
      This value is set to False by default to performance. For recuring use in scripts its recommended to be called at least once at in the script with this value set to True
    .PARAMETER MaxLogFilesToKeep
      Used if ManageLogFiles is True.  Specifies the Maximum number of log files to keep.
    .PARAMETER MaxLogFileSize
      Used if ManageLogFiles is True.  Specifies the Maximum size of a log file before starting a new one.
    .PARAMETER DisplayText
      If True, the passed Text (argument) is written to the display.
    .PARAMETER LogText
      If True, the passed Text (argument) is written to the log file.
    .PARAMETER ReturnLogFile
      If True, the calculated (or passed) log file is returned when the function exits
    .PARAMETER NoNewLine
      If True, No new line is written to the display after the Text (argument) is written.
    .PARAMETER Color
      Color of the String to Log.
    .EXAMPLE
    C:\PS> Write-DMLog -Text "Hey There"
    Hey There

    Note: 
       Debugging of Write-DMLog() can be enabled by setting $DM_DbgLog=$True       

    .EXAMPLE
    C:\PS> Write-DMLog -Text "Hey There" -ReturnLogFile
    Hey There
    C:\Users\DSitner\AppData\Local\Temp\Write-DMLog.log

    Note: 
       1) -ReturnLogfile is passed (and assumed to be $TRUE if not passed) so the calculated log file is returned
       2) $env:Temp is used as the root folder of the logfile because Write-DMLog() is called from the interactive commandline

    .EXAMPLE
    PS C:\Users\DSitner> C:\D\PowerShell\_DMPSLib\Write-DMLogLaunchTest.ps1
    Hey There from a script
    C:\D\PowerShell\_DMPSLib\Logs\Write-DMLogLaunchTest.log

    Note: 
       Write-DMLog() was called from a script and the log file is created in a LOGS subfolder of the script directory, and uses the same filename as the script but usings a .log extension.

    .EXAMPLE
    PS C:\Users\DSitner> get-content (Write-DMLog -Text "Hey There" -ReturnLogFile)
    Hey There
  5/15/2014 11:03:24 AM,Who I am: {ms\dsitner}
    5/15/2014 4:02:33 PM,Hey there
    5/20/2014 4:56:30 PM,hey
    5/20/2014 4:56:50 PM,hey
    10/1/2014 9:01:54 PM,Hey There
    10/1/2014 9:09:43 PM,Hey There

    Note: 
       1) -ReturnLogfile is passed (and assumed to be $TRUE if not passed) so the calculated log file is returned, and then consumed by the get-content commandlet which displays the contents of the log file.
       2) All log file entries are preceeded by a date/time stamp followed by a comma.

    .EXAMPLE
    Scripting Shell Template
    #################################################################################################################
    # ScriptName.ps1
    # Description:
    # Author: David Sitner
    # Version: 1.0.0
    # Repository Path: http://svn.uhc.com/svn/uhgit_drm_des_ls/Powershell/trunk/_DMPSLib/DM_PSIncludeLib_Core.ps1
    #################################################################################################################
    #Include Functions
    $DMToolsDir="\\DES-LS-Dev\NETLOGON\Powershell\_DMPSLib"
    .$DMToolsDir\DM_PSIncludeLib_Core.ps1

    #########################################################
    #Start Of Main()
    #########################################################
    $ScriptDir = (Get-DMScriptDir)
    $ScriptName = (Get-DMScriptName)

    Write-DMLog -Text "Main(): ###################################################################" -CacheLogFileName $True -ManageLogFiles
    Write-DMLog -Text "Main(): ############ Starting Run of $ScriptName"
    Write-DMLog -Text "Main(): ###################################################################"
    Write-DMLog -Text "Main(): ##### PC: {$env:Computername} ######## User: {$env:Username} #######"

    # Do Something

    Write-DMLog -Text "Main(): ###################################################################"
    Write-DMLog -Text "Main(): ############ Ending Run of $ScriptName"
    Write-DMLog -Text "Main(): ###################################################################" -CacheLogFileName $False

    Note: 
       1) 
  #>
  [CmdLetBinding(DefaultParameterSetName = "ManageLogFiles", SupportsShouldProcess=$False)]
  param (
    [Parameter(Mandatory = $True, Position = 0, ValueFromPipeline = $true)]
    [AllowEmptyString()]
     [String]$Text,
    [Parameter(Mandatory = $False, Position = 1)]
     [String]$LogFile,
    [Parameter(Mandatory = $False)]
     [Int]$CacheLogFileName = -1,
    [Parameter(Mandatory = $False)]
     [Switch]$ManageLogFiles = $False,

    [Parameter(Mandatory = $False, ParameterSetName = "UseNumberedLogFiles")]
     [Int]$MaxLogFilesToKeep = 3,
    [Parameter(Mandatory = $False, ParameterSetName = "UseNumberedLogFiles")]
     [Int]$MaxLogFileSize = 100kb,

    [Parameter(Mandatory = $False, ParameterSetName = "UseDateBasedLogFiles")]
    [ValidateSet("LogFilesInSingleFolder","LogFilesInYearSubFolder")]
     [String]$UseDateBasedLogFiles = "",
    [Parameter(Mandatory = $False, ParameterSetName = "UseDateBasedLogFiles")]
     [String]$DateBasedLogFileFormat = "yyyyMMddhhmmtt",
    [Parameter(Mandatory = $False, ParameterSetName = "UseDateBasedLogFiles")]
     [Int]$DateBasedLogFileAgeInDays = 365,
    [Parameter(Mandatory = $False)]
     [Int]$MaxLogTries = 10,
    [Parameter(Mandatory = $False)]
     [Switch]$DisplayText = $True,
    [Parameter(Mandatory = $False)]
     [Switch]$LogText=$True,
    [Parameter(Mandatory = $False)]
     [Switch]$ReturnLogFile = $False,
    [Parameter(Mandatory = $False)]
     [Switch]$NoNewLine = $False
   )

  begin {
    #Check If PS Version is less than 3
    If ($PSVersionTable.PSVersion.Major -lt 3) {
      #If optional Params have no Value - set to default values
      If (!(Test-Path Variable:Private:CacheLogFileName))          {[Int]$CacheLogFileName = -1}
      If (!(Test-Path Variable:Private:ManageLogFiles))            {[Switch]$ManageLogFiles = $False}
      If (!(Test-Path Variable:Private:MaxLogFilesToKeep))         {[Int]$MaxLogFilesToKeep = 3}
      If (!(Test-Path Variable:Private:DateBasedLogFileFormat))    {[Int]$MaxLogFileSize = 100kb}
      If (!(Test-Path Variable:Private:UseDateBasedLogFiles))      {[string]$UseDateBasedLogFiles = ""}
      If (!(Test-Path Variable:Private:DateBasedLogFileFormat))    {[string]$DateBasedLogFileFormat = "yyyyMMddhhmmtt"}
      If (!(Test-Path Variable:Private:DateBasedLogFileAgeInDays)) {[Int]$DateBasedLogFileAgeInDays = 365}
      If (!(Test-Path Variable:Private:MaxLogTries))               {[Int]$MaxLogTries = 10}
      If (!(Test-Path Variable:Private:DisplayText))               {[Switch]$DisplayText = $True}
      If (!(Test-Path Variable:Private:LogText))                   {[Switch]$LogText = $True}
      If (!(Test-Path Variable:Private:ReturnLogFile))             {[Switch]$ReturnLogFile = $False}
      If (!(Test-Path Variable:Private:NoNewLine))                 {[Switch]$NoNewLine = $False}
     }
    #If Debugging, display all pertinant Args and Vars
    If ($DM_DbgLog) {
      Write-Host "Write-DMLog()-Global_WriteDMLog_LogFile:{$Global_WriteDMLog_LogFile}"
      Write-Host "Write-DMLog()-LogFile:{$LogFile}"
      Write-Host "Write-DMLog()-CacheLogFileName:{$CacheLogFileName}"
      Write-Host "Write-DMLog()-LogText:{$LogText}"
      Write-Host "Write-DMLog()-ManageLogFiles:{$ManageLogFiles}"
      Write-Host "Write-DMLog()-MaxLogFilesToKeep:{$MaxLogFilesToKeep}"
      Write-Host "Write-DMLog()-MaxLogFileSize:{$MaxLogFileSize}"
      Write-Host "Write-DMLog()-UseDateBasedLogFiles:{$UseDateBasedLogFiles}"
      Write-Host "Write-DMLog()-DateBasedLogFileFormat:{$DateBasedLogFileFormat}"
      Write-Host "Write-DMLog()-DateBasedLogFileAgeInDays:{$DateBasedLogFileAgeInDays}"
     # Write-Host "Write-DMLog()-:{$}"
     }
   }

  process {  
    #Display the Text
    If ($DisplayText) {
      Write-Host $Text -NoNewline
      }
    If (-Not $NoNewLine) {
      Write-Host
     }

    #If Logging Text (passed Arg)
    If ($LogText) {
      #If $Global_WriteDMLog_LogFile is already defined and its Path -IsValid and No specific Logfile name was passed and $CacheLogFileName <> 1
      If ($Global_WriteDMLog_LogFile -and (Test-Path -LiteralPath $Global_WriteDMLog_LogFile -IsValid) -and ($LogFile -eq "") -and ($CacheLogFileName -ne 1)) {
        #We don't need to calculate $LogFile, just use $Global_WriteDMLog_LogFile
        $LogFile = $Global_WriteDMLog_LogFile
        If ($DM_DbgLog) {Write-Host "Write-DMLog()-Using Cached LogFile=Global_WriteDMLog_LogFile:{$LogFile}"}
       } Else { #We need to calculate $LogFile
        #Else we need to calculate the $LogFile
        If ($DM_DbgLog) {Write-Host "Write-DMLog()-Calculating LogFile"}
        #Test to see arg was passed requesting the Use of Date-Based LogFiles
        if ($UseDateBasedLogFiles -ne "") {
          #If Posh is running in a script
          if (((Get-PSCallStack)[-1]).scriptname) {
            switch ($UseDateBasedLogFiles) {
              #If 
              "LogFilesInSingleFolder" {
                #Log to <ScriptDir>\Logs subfolder and use same name as script but replace extention with .log
                $LogFile = (Get-DMScriptDir) + "\Logs\" + (Get-DMScriptName) + "-" + (get-date -Format $DateBasedLogFileFormat) + ".log"
               }
              #If
              "LogFilesInYearSubFolder" {
                #Log to same folder and name as script but replace extention with .log
                $LogFile = (Get-DMScriptDir) + "\Logs\" + (get-date -Format yyyy) + "\" + (Get-DMScriptName) + "-" + (get-date -Format $DateBasedLogFileFormat)  + ".log"
               }
             }
            If ($DM_DbgLog) {Write-Host "Write-DMLog()-Posh is running in a script. LogFile: $LogFile"}
           } Else {
            #Log to <ScriptDir>\Logs subfolder and use same name as script but replace extention with .log
            $LogFile = (Get-DMScriptDir) + "\Write-DMLog-" + (get-date -Format $DateBasedLogFileFormat) + ".log"
            If ($DM_DbgLog) {Write-Host "Write-DMLog()-Posh is not running in a script. Calculated LogFile: $LogFile"}
           }
         } Else { #Not using Date-Based LogFiles
          #If No LogFile was passed
          if ($LogFile -eq "") {
            #If Posh is not running in a script
            if (-not([bool](((Get-PSCallStack)[-1]).scriptname))) {
              #Log to fixed name log file in Temp dir
              $LogFile = (Get-DMScriptDir) + "\Write-DMLog.log"
              If ($DM_DbgLog) {Write-Host "Write-DMLog()-Posh is not running in a script. Calculated LogFile: $LogFile"}
             } Else {
              #Log to same folder and name as script but replace extention with .log
              $LogFile = (Get-DMScriptDir) + "\Logs\" + (Get-DMScriptName) + ".log"
              If ($DM_DbgLog) {Write-Host "Write-DMLog()-Posh is running in a script. LogFile: $LogFile"}
             }
           } Else { #LogFile Was passed
            #Parse Passed LogFile
            $LogFile = $LogFile.Trim()
            #If Passed Logfile is an absolute path use it as is
            If (Split-Path -IsAbsolute -Path $LogFile){
              $LogFile = $LogFile  
              If ($DM_DbgLog) {Write-Host "Write-DMLog()-Absolute LogFile path was passed. LogFile:{$LogFile}"}
             #Else Create an absolute path from passed filename
             } Else {
              $LogFile = Join-Path -Path "$(Get-DMScriptDir)\Logs\" -ChildPath $LogFile
              If ($DM_DbgLog) {Write-Host "Write-DMLog()-Relative LogFile path was passed. Calulated LogFile:{$LogFile}"}
             }
           } #LogFile Was passed
         } #using Date-Based LogFiles
       } #We need to calculate $LogFile

      #Write Text to Log using TryCatch block in case the folder doesn't exist - Try {Write} Catch {Create Folder;Write}
      $blnLogDone = $false
      $NumLogTries = 0
      do {
        try {
          add-content -path $LogFile -Value $((get-date).tostring() + ","+$Text) -Force -ErrorAction Stop
          $blnLogDone = $true
         } Catch {
          $NumLogTries ++
          #If logfile directory does not exist
          If (-Not(Test-Path -Path (Split-Path $LogFile))) {
            #Create directory for LogFileSplit-Path $LogFile
            New-Item -Path (Split-Path $LogFile) -ItemType directory -ErrorAction SilentlyContinue | Out-Null
           }
          If ($NumLogTries -lt $MaxLogTries) {
            If ($DM_DbgLog) {Write-host "NumLogTries: {$NumLogTries}"}
            Start-Sleep -Milliseconds (Get-Random -Maximum 100)
           } Else {
            Write-host "Warning: Reached MaxLogTries: {$MaxLogTries} attempts to write to logfile file : {$LogFile}. Aborting affort to write to file."
           }
         }    
       } while ((-not $blnLogDone) -and ($NumLogTries -lt $MaxLogTries))
     } Else {# Not Logging Text
      If ($DM_DbgLog) {Write-Host "Write-DMLog()-Not Logging"}
     }
   }#process

  #LogFile Pruning process
  end {
    #Manage caching of the logfile for subsequent calls to Write-DMLog
    switch ($CacheLogFileName) {
      #If 1 CacheLogFileName is True create global variable Global_WriteDMLog_LogFile
      1 {

        #If $Global_WriteDMLog_LogFile exists
        If ($Global_WriteDMLog_LogFile) {
          #If it doesn't already match $LogFile
          If ($Global_WriteDMLog_LogFile -ne $LogFile) {
            $Global_WriteDMLog_LogFile = $LogFile
            If ($DM_DbgLog) {Write-Host "Write-DMLog()-Reset existing Global_WriteDMLog_LogFile to:{$Global_WriteDMLog_LogFile}"}
           }
         } Else {
          New-Variable -Name Global_WriteDMLog_LogFile -Value $LogFile -Scope "Global" -Option ReadOnly
          If ($DM_DbgLog) {Write-Host "Write-DMLog()-Caching the logfile for subsequent calls via Global var Global_WriteDMLog_LogFile:{$Global_WriteDMLog_LogFile}"}
         }
       }
      #If 0 CacheLogFileName is False - Remove global variable Global_WriteDMLog_LogFile
      0 {
        Remove-Variable Global_WriteDMLog_LogFile -Scope "Global" -Force -ErrorAction SilentlyContinue
        If ($DM_DbgLog) {Write-Host "Write-DMLog()-Removing Global Logfile Cache var Global_WriteDMLog_LogFile:{$Global_WriteDMLog_LogFile}"}
       }
      #If -1 CacheLogFileName is Undefined (Not passed)
      Default {
        If ($DM_DbgLog) {Write-Host "Write-DMLog()-CacheLogFileName is in unconfigured state:{$CacheLogFileName}"}
       }
     }

    #If Managing LogFiles
    If ($ManageLogFiles) {
      if ($UseDateBasedLogFiles -ne "") {
        $DateThreshold = (get-date) - (New-TimeSpan -days $DateBasedLogFileAgeInDays)
        Get-ChildItem -Path (Split-Path -Path $LogFile -Parent) | Where-Object {$_.LastWriteTime -lt $DateThreshold} | Remove-Item -Force -ErrorAction SilentlyContinue
       } Else { 
        #If LogFile length > MaxLogFileSize
        if ((Get-ChildItem $LogFile).Length -gt $MaxLogFileSize) {
          $LogFile_ = ($LogFile -split "\.")[0] + "_"
          #$LogFile_ => H:\PowerShell\Write-DMLog_
          If (Test-Path $($LogFile_ + ($MaxLogFilesToKeep).ToString() + ".log")) {
            Remove-item -Path $($LogFile_ + ($MaxLogFilesToKeep).ToString() + ".log")
           }
          for($i=$MaxLogFilesToKeep-1;$i -ge 1;$i--) {
            If (Test-Path $($LogFile_ + ($i).ToString() + ".log")) {
              Rename-item -Path $($LogFile_ + ($i).ToString() + ".log") -NewName $($LogFile_ + ($i+1).ToString() + ".log")
             }
           }
          Rename-item -Path $LogFile -NewName $($LogFile_ + "1.log")
         }
       } # UseDateBasedLogFiles
     } #ManageLogFiles
    If ($ReturnLogFile) {return $LogFile}
   }#end

 } #End Function Write-DMLog

Function Invoke-DMMultiThreadingEngine {
    <#
    .SYNOPSIS
      Writes passed string to display console and to a managed log file in a way that is intended to be flexible, powerful and lightweight.
    .DESCRIPTION
      PowerShell script logging function with usable default options for most cases.  Provides flexibility to modifying defaults through optional arguments.
      Note: 
       Debugging of Write-DMLog() can be enabled by setting $DM_DbgLog=$True
    .PARAMETER Text
      Text string to write to console display and to the log file.
    .PARAMETER LogFile
      Path to the log file
      Options: If passed LogFile is a valid absolute path, it is used as passed.
        If not, it is interpreted as a -childpath to the Join-Path commandlet, using a root folder of <ScriptDir>\Logs - where <ScriptDir> is the folder containing the running script.
        If the Write-DMLog is called from the interactive commandline, <ScriptDir> is assigned the $env:Temp folder (by DM Library function Get-DMScriptDir())
    .PARAMETER CacheLogFileName
      CacheLogFileName is an integer which is used to represent 3 states: True, False and UnDefined(default).
      It does this by being set to 1 of 3 values: 1=True, 0=False and -1=Undefined
       If set to True via Syntax: -CacheLogFileName $True
          the LogFile name is cached and and saved in the the global variable: $Global_WriteDMLog_LogFile
       If set to Undefined(default value if parameter is not specified)
          the cached LogFile name is used if $Global_WriteDMLog_LogFile has been set previously, otherwise the logfile is calculated 
      Note: 
       This is implemented as a TriState integer so that True/False logic can be used when passing and scripting with the parameter.
       This technique works because [boolean]1 = [boolean]-1 = True, and [boolean]0 = False
       So CacheLogFileName
    .PARAMETER UseCacheLogFileName
      To increase performance the LogFile name is cached and not recalculated each time the function is called.
      However if within a script multiple log files are to be written to, use this switch -CacheLogFileName:$false  
    .PARAMETER UseDateBasedLogFiles
      Overrides the default naming convention for Logfiles and uses a DateBased LogFile format of:
      <ScriptDir>\Logs\<ScriptName>-yyyyMMddhhmmtt
      For Example: C:\ScriptDir\Logs\ScriptName-201410020639PM.log
    .PARAMETER DateBasedLogFileFormat
      Overrides the default DateBased LogFile format of <ScriptDir>\Logs\<ScriptName>-yyyyMMddhhmmtt
      For options see: http://msdn.microsoft.com/en-us/library/system.globalization.datetimeformatinfo(VS.85).aspx
    .PARAMETER DateBasedLogFileAgeInDays
      The age threshold (in days) of the oldest log file to keep.
      This parameter is only considered if called at the same time as the ManageLogFiles is passed as True.
      If ManagerLogFiles is Ture, older log files will be deleted.
    .PARAMETER ManageLogFiles
      If True, log files will be pruned to keep file size below MaxLogFileSize (argument), previous log files will be renamed by appending an _# to the name, and deleted to keep the log files under MaxLogFilesToKeep (argument).
      This value is set to False by default to performance. For recuring use in scripts its recommended to be called at least once at in the script with this value set to True
    .PARAMETER MaxLogFilesToKeep
      Used if ManageLogFiles is True.  Specifies the Maximum number of log files to keep.
    .PARAMETER MaxLogFileSize
      Used if ManageLogFiles is True.  Specifies the Maximum size of a log file before starting a new one.
    .PARAMETER DisplayText
      If True, the passed Text (argument) is written to the display.
    .PARAMETER LogText
      If True, the passed Text (argument) is written to the log file.
    .PARAMETER ReturnLogFile
      If True, the calculated (or passed) log file is returned when the function exits
    .PARAMETER NoNewLine
      If True, No new line is written to the display after the Text (argument) is written.
    .PARAMETER Color
      Color of the String to Log.
    .EXAMPLE
    C:\PS> Write-DMLog -Text "Hey There"
    Hey There

    Note: 
       Debugging of Write-DMLog() can be enabled by setting $DM_DbgLog=$True       

    .EXAMPLE
    C:\PS> Write-DMLog -Text "Hey There" -ReturnLogFile
    Hey There
    C:\Users\DSitner\AppData\Local\Temp\Write-DMLog.log

    Note: 
       1) -ReturnLogfile is passed (and assumed to be $TRUE if not passed) so the calculated log file is returned
       2) $env:Temp is used as the root folder of the logfile because Write-DMLog() is called from the interactive commandline

    .EXAMPLE
    PS C:\Users\DSitner> C:\D\PowerShell\_DMPSLib\Write-DMLogLaunchTest.ps1
    Hey There from a script
    C:\D\PowerShell\_DMPSLib\Logs\Write-DMLogLaunchTest.log

    Note: 
       Write-DMLog() was called from a script and the log file is created in a LOGS subfolder of the script directory, and uses the same filename as the script but usings a .log extension.

    .EXAMPLE
    PS C:\Users\DSitner> get-content (Write-DMLog -Text "Hey There" -ReturnLogFile)
    Hey There
  5/15/2014 11:03:24 AM,Who I am: {ms\dsitner}
    5/15/2014 4:02:33 PM,Hey there
    5/20/2014 4:56:30 PM,hey
    5/20/2014 4:56:50 PM,hey
    10/1/2014 9:01:54 PM,Hey There
    10/1/2014 9:09:43 PM,Hey There

    Note: 
       1) -ReturnLogfile is passed (and assumed to be $TRUE if not passed) so the calculated log file is returned, and then consumed by the get-content commandlet which displays the contents of the log file.
       2) All log file entries are preceeded by a date/time stamp followed by a comma.

    .EXAMPLE
    Scripting Shell Template
    #################################################################################################################
    # ScriptName.ps1
    # Description:
    # Author: David Sitner
    # Version: 1.0.0
    # Repository Path: http://svn.uhc.com/svn/uhgit_drm_des_ls/Powershell/trunk/_DMPSLib/DM_PSIncludeLib_Core.ps1
    #################################################################################################################
    #Include Functions
    $DMToolsDir="\\DES-LS-Dev\NETLOGON\Powershell\_DMPSLib"
    .$DMToolsDir\DM_PSIncludeLib_Core.ps1

    #########################################################
    #Start Of Main()
    #########################################################
    $ScriptDir = (Get-DMScriptDir)
    $ScriptName = (Get-DMScriptName)

    Write-DMLog -Text "Main(): ###################################################################" -CacheLogFileName $True -ManageLogFiles
    Write-DMLog -Text "Main(): ############ Starting Run of $ScriptName"
    Write-DMLog -Text "Main(): ###################################################################"
    Write-DMLog -Text "Main(): ##### PC: {$env:Computername} ######## User: {$env:Username} #######"

    # Do Something

    Write-DMLog -Text "Main(): ###################################################################"
    Write-DMLog -Text "Main(): ############ Ending Run of $ScriptName"
    Write-DMLog -Text "Main(): ###################################################################" -CacheLogFileName $False

    Note: 
       1) 
  #>
 [CmdLetBinding(SupportsShouldProcess=$False)]
 param (
   [Parameter(Mandatory = $True, Position = 0)]
    [AllowEmptyCollection()]
    [Array]$RecordArray,
   [Parameter(Mandatory = $True, Position = 1)]
    [ScriptBlock]$MTScriptBlock,
   [Parameter(Mandatory = $False)]
    [Switch]$RunJobsOnRemoteComputers = $False,
   [Parameter(Mandatory = $False, Position = 3)]
    [Int]$MaxThreads=10,
   [Parameter(Mandatory = $False)]
    [Int]$MaxJobFailuresPerRecordAllowed=1,
   [Parameter(Mandatory = $False)]
    [Int]$HungJobThresholdinSeconds=600,
   [Parameter(Mandatory = $False)]
    [Switch]$AllowScriptBlockDebugging = $False,
   [Parameter(Mandatory = $False)]
    [Int]$BatchSaveSize = 0,
   [Parameter(Mandatory = $False)]
    [Int]$AgeInDaysOfOldResultsFoldersToDelete = 30,
   [Parameter(Mandatory = $False)]
    [Array]$SBArgArray,
   [Parameter(Mandatory = $False)]
    [String]$ResultsDir = [string](Get-DMScriptDir)+"\Results",
   [Parameter(Mandatory = $False)]
    [String]$NameQualifier = [string](Get-DMScriptName),
   [Parameter(Mandatory = $False)]
    [PSCredential]$Credential
  )

 Function CheckFunctionPreReqs {
  param ([PSCredential]$Cred)
   $intErrCode = 0
   $intErrMsg = ""
   
   #Use Internal function to confirm we're running in Admin mode
   If (-Not (Is-DMAdmin)) {
       $intErrCode = 101
       $strErrMsg = "Fatal Error: This function needs to be run in Elevated Powershell host console.  Terminating Function."
      
    } Else {
     #If $Cred arg was passed check if it's valid using the Confirm-DMCred function
     If ($Cred -and (-not(Confirm-DMCred($Cred)))) {
       $intErrCode = 102
       $strErrMsg = "Fatal Error: Invalid Credential Passed in -Credential Argument.  Terminating Function."
      }
    }
   #Display and Error Messages
   If ($intErrCode -ne 0) {
	 Write-Error -Message $strErrMsg
    }

   $objChkPreReqsResults = new-object PSObject
   $objChkPreReqsResults | add-member -membertype NoteProperty -name "ReturnedErrorCode" -Value $intErrCode
   $objChkPreReqsResults | add-member -membertype NoteProperty -name "ReturnedErrorMsg" -Value $strErrMsg
   Return $objChkPreReqsResults
  }
 
 Function Set-DMMaxPSSessions {
  param ([int]$MaxThreads)
   #Collect Original WSMan Configuration Values required to remove the limit on the number of PSSessions
   $Shell_MaxShellsPerUser_Orig = [Int](get-item wsman:\localhost\Shell\MaxShellsPerUser).value
   $Quotas_MaxShellsPerUser_Orig = [Int](get-item wsman:\localhost\Plugin\microsoft.powershell\Quotas\MaxShellsPerUser).value
   $Quotas_MaxShells_Orig = [Int](get-item wsman:\localhost\Plugin\microsoft.powershell\Quotas\MaxShells).value
   $blnWSManChanged=$False
   #If needed, configure WSMan Values required to allow the number of PSSessions to match $MaxThreads
   If ($Shell_MaxShellsPerUser_Orig -lt $MaxThreads) {
     set-item wsman:\localhost\Shell\MaxShellsPerUser $MaxThreads
     $blnWSManChanged=$true
    }
   If ($Quotas_MaxShellsPerUser_Orig -lt $MaxThreads) {
     set-item wsman:\localhost\Plugin\microsoft.powershell\Quotas\MaxShells $MaxThreads
     $blnWSManChanged=$true
    }
   If ($Quotas_MaxShells_Orig -lt $MaxThreads) {
     set-item wsman:\localhost\Plugin\microsoft.powershell\Quotas\MaxShellsPerUser $MaxThreads
     $blnWSManChanged=$true
    }
   If ($blnWSManChanged) {
     Restart-Service winrm
    }
  }

 Function MTWriteProgress {
  param ($MTStartTime,$MTJobPrefix,$MTSessionPrefix,$MaxThreads,$RcdNum,$RcdArrayCount,$AllowScriptBlockDebugging)
  $JobCount=(get-job -Name "$MTJobPrefix*").count
  $JobCompletedCount = (get-job -Name "$MTJobPrefix*" | where {$_.state -eq "Completed"}).count
  $SessionCount = (Get-PSSession -Name "$MTSessionPrefix*").count
  $Now = Get-Date
  $ElapsedMins = ($Now - $MTStartTime).totalminutes
  $ProcessingRate = $RcdNum / $ElapsedMins
  $EstMinsRemaining = ($RcdArrayCount - $RcdNum) / $ProcessingRate
  $EstFinishTime = [DateTime]::Now.AddMinutes($EstMinsRemaining)
  $PercentComplete = ($RcdNum/$RcdArrayCount)
  If ($PercentComplete -gt 1) {
    If ($RcdNum -gt ($RcdArrayCount+1)) {
      Write-DMLog -Text ("Get-DMScriptName (): Warning: RcdNum exceeded more than expected max of MaxRecordNum + 1. RcdNum: {$RcdNum}. MaxRecordNum: {$MaxRecordNum}")
     }
    $PercentComplete = 1
   }
  If ($AllowScriptBlockDebugging) {
    $MaxThreadsDisplayText = "[ScriptBlockDebuggingEnabled]"
   } else {
    $MaxThreadsDisplayText = "MaxThreads: $MaxThreads."
   }
  $Progress_Activity = "$MaxThreadsDisplayText   SessionCount: $SessionCount.   Running background jobs: $JobCount.  Completed jobs: $JobCompletedCount.`
    Percent Complete: {0:P1}.      Active Fail Records: {1:N0}.   Fixed Fail Records: {2:N0}.   Fail Rcds Exceeding Retry Count: {3:N0}.   Number of Hung Jobs Killed: {4:N0}" -f $PercentComplete,$hshFailingJobsTracker.Count,$hshFailedThenSuceededRcds.count,$arrFailedRcds.count,$arrHungRcds.count
  $Progress_Status = "Processed $RcdNum Records out of a Maximum of $RcdArrayCount.  Processing Rate {0:N1} (Records/minute).`
    Elapsed Minutes: {1:N1}.  Estimated Minutes remaining {2:N1}.  Estimated completion time $EstFinishTime" -f $ProcessingRate,$ElapsedMins,$EstMinsRemaining
  Write-Progress -Activity $Progress_Activity -status $Progress_Status -percentComplete ($PercentComplete*100)
 }

 Function ManageOldBatchResults {
  param ($ResultsDir, $NameQualifier, $AgeInDaysOfOldResultsFoldersToDelete)
  #Create $ResultsDir if not already there
  If (-not (test-path -Path $ResultsDir)) {
    Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): Creating ResultsDir: {$ResultsDir}"
    New-Item -Path $ResultsDir -ItemType directory -ErrorAction SilentlyContinue | Out-Null
   }
  #Define Name Qualifier for renaming Archived ResultsDirs
  $ArchivedResultsNameQualifier = "ArchivedResults"
  #Get the list of Old Batch Files In the ResultsDir, sort by date
  $OldBatchFilesInResultsDir = Get-ChildItem "$ResultsDir\$NameQualifier*.xml" | sort -Property LastWriteTime
  If ($OldBatchFilesInResultsDir.count -gt 0) {
    #Use the datetime of the first batch file to create a unique time-stamped subfolder to move the files into
    # Sample: E:\Scripts\MaintainLSTrackingFiles\Results\ArchivedResults-MaintainLSTrackingFiles-20150313090334PM
    $SubDirForOldBatchFiles = "$ResultsDir\$ArchivedResultsNameQualifier-$NameQualifier-$(get-date -Date  ($OldBatchFilesInResultsDir[0]).LastWriteTime -format "yyyyMMddhhmmsstt")"
    #Log File Move
    Write-DMLog -Text "Invoke-DMMultiThreadingEngine.ManageOldBatchResults() - Moving {$(($OldBatchFilesInResultsDir).count)} previously existing results files to Archive folder: {$SubDirForOldBatchFiles}"
    #Create Folder if it does not exist
    If (-not (test-path -Path $SubDirForOldBatchFiles)) {New-Item -Path $SubDirForOldBatchFiles -ItemType directory -ErrorAction SilentlyContinue | Out-Null}
    #Move OldBatchFilesInResultsDir to the newly created SubDirForOldBatchFiles
    $OldBatchFilesInResultsDir | Move-item -Destination $SubDirForOldBatchFiles
   }
  #Delete Old Results Folders older than $AgeInDaysOfOldResultsFoldersToDelete days
  $arrOldResultsFoldersToDelete = Get-ChildItem -Path $ResultsDir -Directory |
                                  where {($_.CreationTime -lt ((Get-Date)-(New-TimeSpan -days $AgeInDaysOfOldResultsFoldersToDelete)))}
  foreach ($OldResultsFoldersToDelete in $arrOldResultsFoldersToDelete) {
    Write-DMLog -Text "Invoke-DMMultiThreadingEngine.ManageOldBatchResults() - Deleting Old Archive Folder: {$(($OldResultsFoldersToDelete).Fullname)}"
    $OldResultsFoldersToDelete | Remove-Item -Recurse -Force
   }
 }

 #Initialize Vars
 $arrobjResults = @()              #Initialize array of Result objects to hold the returned objects from the script block jobs.
                                    #Array is returned back to the calling script as part of $objReturnResult
 $arrFailedRcds = @()              #Array of Records which failed to process correctly after failing $MaxJobFailuresPerRecordAllowed times.
                                    #Array is returned back to the calling script as part of $objReturnResult
 $arrObjJobFailureErrors = @()     #Array of objects containing 4 fields: Rcd,ErrorCode,ErrMsg,TimeStamp
                                    #Array is returned back to the calling script as part of $objReturnResult
 $arrHungRcds = @()                #Array of Records whose associated jobs ran longer than $HungJobThresholdinSeconds seconds.
                                    #Array is returned back to the calling script as part of $objReturnResult
 $hshFailedThenSuceededRcds = @{}  #Name:RcdNum = Value:FailureCount (Represents the number of times the job for this RcdNum failed before succeeding)
                                    #Hash table is returned back to the calling script as part of $objReturnResult
 $hshFailingJobsTracker = @{}      #Name:RcdNum = Value:FailureCount - Hash to track dynamic growth and shrinkage of failed Records via the RcdNum in $RecordArray

 $MTJobPrefix = "Job_$NameQualifier-"
 $MTSessionPrefix = "Session_$NameQualifier-"
 $MaxRecordNum = $RecordArray.count
 $SBArgArrayToBeCalculatedInJobLoop = @('$Rcd')
 $NoRecordSelectedPlaceholder = [guid]::NewGuid()    #Unique value unlikely to be found in passed $RecordArray
 $RcdNum = 0
 $JobNum = 0
 $SessionNum = 0
 $BatchSaveCount = 0
 $LastErrorCount = 0
 $ReturnedErrorCode = 0
 $ReturnedErrorMsg = ""
 $MTStartTime = Get-Date
 If ($BatchSaveSize -gt 0){
   $blnSavingResultsToXML = $true
  } else {
   $blnSavingResultsToXML = $false
  }
 
 $MTLogFile = Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): ########################################################################" -ReturnLogFile
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): MTStartTime:     {$MTStartTime}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): Starting Invoke-DMMultiThreadingEngine() with the following Parameters:"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Log File :     {$MTLogFile}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   MaxRecordNum:  {$MaxRecordNum}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   ResultsDir:    {$ResultsDir}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   NameQualifier: {$NameQualifier}"
 If ($AllowScriptBlockDebugging) {
   Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   MaxThreads:    {1} <-- [ScriptBlockDebuggingEnabled]"
  } else {
   Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   MaxThreads:    {$MaxThreads}"
  }
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   BatchSaveSize: {$BatchSaveSize}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   RunJobsOnRemoteComputers:       {$RunJobsOnRemoteComputers}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   HungJobThresholdinSeconds:      {$HungJobThresholdinSeconds}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   MaxJobFailuresPerRecordAllowed: {$MaxJobFailuresPerRecordAllowed}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): ########################################################################"

 #Check Function PreReqs
 $objPreReqsChkResults = CheckFunctionPreReqs ($Credential)
 $ReturnedErrorCode = $objPreReqsChkResults.ReturnedErrorCode
 $ReturnedErrorMsg = $objPreReqsChkResults.ReturnedErrorMsg
 
 #Modify as needed the WSMan paramters to allow $maxthreads PSSessions
 Set-DMMaxPSSessions ($Maxthreads)

 #Manage Old Batch Results
 ManageOldBatchResults $ResultsDir $NameQualifier $AgeInDaysOfOldResultsFoldersToDelete

 #Clean out any old Jobs and PSSessions
 Get-Job -Name "$MTJobPrefix*" | Remove-Job -Force
 Get-PSSession -Name "$MTSessionPrefix*" | Remove-PSSession
 
 #Loop through each record until PreReqsChk failed, were out of New Rcds, done retrying failed records, and we're out of jobs
 While (($objPreReqsChkResults.ReturnedErrorCode -eq 0) -and ($MaxRecordNum -gt 0) -and (($RcdNum -lt $MaxRecordNum) -or ($arrFailedRcdNumsNotInProgress.Count -gt 0) -or (((get-job -Name "$MTJobPrefix*").count -gt 0) -or ($RcdNum -eq 0)))) {
   #Initialize $Rcd to a GUID value that can't be confused with a value passed in the array
   $Rcd = $NoRecordSelectedPlaceholder

   #While (we have more Records to process) and (we haven't reached our $MaxThreads) and (We're NOT Allowing ScriptBlock Debugging)
   While ((($RcdNum -lt $MaxRecordNum) -or (([array]($hshFailingJobsTracker.keys) | select -unique | where {$MTJobs.rcdnum -notcontains $_}).count -gt 0)) -and ((get-job -Name "$MTJobPrefix*").count -lt $MaxThreads) -and -not $AllowScriptBlockDebugging) {
     #Get array of all jobs matching the specified (or default) naming convention 
     $MTJobs = Get-Job -Name "$MTJobPrefix*"
     #Initialize $Rcd to a GUID value that can't be confused with a value passed in the array
     $Rcd = $NoRecordSelectedPlaceholder

     #Logic below dictates that known failed records will be processed ASAP - before remaing unprocess records in passed $RecordArray
     #If we are tracking failed records that haven't exceeded the $MaxJobFailuresPerRecordAllowed
     $FailedRcdNum = -1
     If ($hshFailingJobsTracker.Count -gt 0) {
       #Get the array of In-Progress Records - We track this using the RcdNum noteproperty we custom add to the Job objects
       $arrInProgressRcdNums = $MTJobs.RcdNum
       #Use Regex to quickly calculate the array of Failed RcdNums Not currently In Progress
       [regex]$arr1_regex = ‘(?i)^(‘ + (($arrInProgressRcdNums |foreach {[regex]::escape($_)}) –join “|”) + ‘)$’
       $arrFailedRcdNumsNotInProgress = ([array]$hshFailingJobsTracker.keys) -notmatch $arr1_regex
       If ($arrFailedRcdNumsNotInProgress.Count -gt 0) {
         #Use Record having the first RcdNum in hash table $hshFailingJobsTracker
         $FailedRcdNum = $arrFailedRcdNumsNotInProgress[0]
         $Rcd = $RecordArray[$FailedRcdNum]
        }
      }
     #If we didn't have any Failed Records to Retry AND we still have Records from $RecordArray
     If (($Rcd -eq $NoRecordSelectedPlaceholder) -AND ($RcdNum -le $MaxRecordNum)){
       #Use next record in the passed array
       $Rcd = $RecordArray[$RcdNum]
       $RcdNum += 1
      }

     #If we found record to process - either a new record from $RecordArray or retrying a failed job record
     If ($Rcd -ne $NoRecordSelectedPlaceholder) {
       #Calculate ArgList to pass to ScriptBlock
       # Call Example: -SBArgArray @($ReadOnly) -SBArgArrayToBeCalculatedInJobLoop '$Rcd','$Color'
       $SBArgArrayCalculated = @($SBArgArrayToBeCalculatedInJobLoop | Invoke-Expression)
       $ArgList = $SBArgArrayCalculated + $SBArgArray

       if ($RunJobsOnRemoteComputers) {
         Remove-variable -Name ComputerName -Force -ErrorAction SilentlyContinue
         #Extract the ComputerName from Record
         # If RunJobsOnRemoteComputers is specified by the calling script
         # The ComputerName can be passed in the multiple ways below:
         # If the $RecordArray is an array of strings, they are assumed to be an array of computer names
         If ($Rcd -is [String]) { $ComputerName = $Rcd
          } Else {
           # If the $RecordArray is an array of objects with a string attribute of "name", that attribute is assumed to contain the computer name
           # Note This supports passing an array of AD computer objects via the Get-ADComputer Cmdlet
           If ($Rcd.name -is [String]) { $ComputerName = $Rcd.name
            } Else {
             # If the $RecordArray is an array of objects with a string attribute of "computer", that attribute is assumed to contain the computer name
             If ($Rcd.Computer -is [String]) { $ComputerName = $Rcd.Computer
              } Else {
               # If the $RecordArray is an array of objects with a string attribute of "PC", that attribute is assumed to contain the computer name
               If ($Rcd.PC -is [String]) { $ComputerName = $Rcd.PC
                } Else {
                 # If the $RecordArray is an array of objects with a string attribute of "Server", that attribute is assumed to contain the computer name
                 If ($Rcd.Server -is [String]) { $ComputerName = $Rcd.Server
                  } Else {
                   # If the $RecordArray is an array of objects with a string attribute of "Host", that attribute is assumed to contain the computer name
                   If ($Rcd.Host -is [String]) { $ComputerName = $Rcd.Host
                  }
                }
              }
            }
          }
        }
       
       $ComputerName = $ComputerName.trim()
       If ($ComputerName) {
         $JobName = $MTJobPrefix + "Job:" + $JobNum + "-PC:" + $ComputerName
         If ([string]::IsNullOrEmpty($Credential)) {
           $Job = Invoke-Command -ComputerName $ComputerName -ScriptBlock $MTScriptBlock -ArgumentList $ArgList -AsJob -JobName $JobName
          } Else {
           $Job = Invoke-Command -ComputerName $ComputerName -ScriptBlock $MTScriptBlock -ArgumentList $ArgList -AsJob -JobName $JobName -Credential $Credential
          }
        }
      } Else { #Not Running Jobs On RemoteHosts
       #Get the array of in-Use SessionNums (Using the SessionNum noteproperty we custom add to the Job objects)
       $arrInUseSessionNums = $MTJobs.SessionNum
       #Use Regex to quickly calculate the array of availables Session Numbers
       [regex]$arr1_regex = ‘(?i)^(‘ + (($arrInUseSessionNums |foreach {[regex]::escape($_)}) –join “|”) + ‘)$’
       $arrAvailableSessions = (1..$MaxThreads) -notmatch $arr1_regex
       #If any Sessions are available for running another job
       If ($arrAvailableSessions.Count -gt 0) {
         $JobNum += 1
         #Use the first available Session Number
         $SessionNum = $arrAvailableSessions[0]
         #Build the Session name
         $SessionName = "$MTSessionPrefix$SessionNum"
         #Get array of existing PS Sessions matching our naming convention
         $arrMTSessions = Get-PSSession -Name "$MTSessionPrefix*"
         #Get the array of SessionNums (Using the SessionNum noteproperty we custom added to our Session objects)
         $arrExistingSessionNums = $arrMTSessions.SessionNum
         #If the selected Session doesn't exist yet
         #IF the SessionNum belongs to a session that has already been created and is currently available to use for the next job
         If ($arrExistingSessionNums -contains $SessionNum) {
           $Session = Get-PSSession -Name $SessionName
          } Else { #Create a New Session with the correct Number
           If ([string]::IsNullOrEmpty($Credential)) {
             $Session = New-PSSession -Name $SessionName -EnableNetworkAccess:$True
            } Else {
             $Session = New-PSSession -Name $SessionName -EnableNetworkAccess:$True -Credential $Credential
            }
           #Add SessionNum as a property to the newly created Session so we can easily determine which of our Sessions we've already created
           $Session | Add-Member -membertype NoteProperty -name "SessionNum" -Value $SessionNum
          }
         #Launch ScriptBlock as a PS Job
         #$Job = Start-Job -ScriptBlock $MTScriptBlock -Name $MTJobPrefix"Job:"$JobNum"-Ses:"$SessionNum -ArgumentList $ArgList
         $JobName = $MTJobPrefix + "Job:" + $JobNum + "-Ses:" + $SessionNum
         $Job = Invoke-Command -ScriptBlock $MTScriptBlock -ArgumentList $ArgList -Session $Session -AsJob -JobName $JobName
         #Add SessionNum as a property to the newly created Job so we can track which Sessions are in use
         $Job | Add-Member -membertype NoteProperty -name "SessionNum" -Value $SessionNum
        }
      }#End Else Not Running Jobs On RemoteHosts
       #Regardless of which host the job is running on,
       # add RcdNum as a property to the newly created Job so we can track which Rcd is being processed by the job in case it fails
       If ($FailedRcdNum -eq -1) {
         $Job | Add-Member -membertype NoteProperty -name "RcdNum" -Value ($RcdNum - 1)
        } Else {
         $Job | Add-Member -membertype NoteProperty -name "RcdNum" -Value ($FailedRcdNum)
        }
      } Else { #No more records left to process from $RecordArray or $hshFailingJobsTracker
       Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): No more records left to process from passed array of records or from failed Jobs to retry"
      }
     #Write-Progress
     MTWriteProgress $MTStartTime $MTJobPrefix $MTSessionPrefix $MaxThreads $RcdNum $MaxRecordNum $AllowScriptBlockDebugging
    }
 
   ##############################
   # Allow ScriptBlock Debugging
   ##############################
   If ($AllowScriptBlockDebugging) {
     $Rcd = $RecordArray[$RcdNum]
     $RcdNum += 1
     #Calculate ArgList to pass to ScriptBlock
     # Call Example: -SBArgArray @($ReadOnly) -SBArgArrayToBeCalculatedInJobLoop '$Rcd','$Color'
     $SBArgArrayCalculated = @($SBArgArrayToBeCalculatedInJobLoop | Invoke-Expression)
     $ArgList = $SBArgArrayCalculated + $SBArgArray
     ####################################################################################
     #            TO DEBUG THE SCRIPTBLOCK WITHIN POWERSHELL_ISE
     # 1) Call Invoke-DMMultiThreadingEngine() with argument: -AllowScriptBlockDebugging
     # 2) Set breakpoint using F9 key line AFTER these comments
     # 3) Run the calling script
     # 4) When breakpoint below is reached execution will stop,
     #    press F11 key to step into and debug the passed ScriptBlock
     ####################################################################################
     $arrobjResults += $MTScriptBlock.Invoke($ArgList)
     #Write-Progress
     MTWriteProgress $MTStartTime $MTJobPrefix $MTSessionPrefix $MaxThreads $RcdNum $MaxRecordNum $AllowScriptBlockDebugging
    } Else {
     ######################################################### 
     #Receive Results from completed Jobs and then remove them
     ######################################################### 
     $CompletedJobs = Get-Job -Name "$MTJobPrefix*" | where {($_.state -eq "Completed")}
     $arrobjResults += $CompletedJobs | Receive-Job
     #Check if any RcdNums from the Completed jobs need to be removed from the $hshFailingJobsTracker hash table
     If ($hshFailingJobsTracker.count -gt 0) {
       #$arrCompletedRcds = $RecordArray | select -Index ([array]($CompletedJobs.RcdNum))
       #Use Regex to quickly calculate the common RcdNums in the Completed array and Failed hash table
       [regex]$arrCompletedRcdNums_Regex = ‘(?i)^(‘ + (($CompletedJobs.RcdNum |foreach {[regex]::escape($_)}) –join “|”) + ‘)$’
       $arrCommonRcdNumsInCompletedandFailed = ([array]$hshFailingJobsTracker.Keys) -match $arrCompletedRcdNums_Regex
       #For each Previously failed RcdNum now completed
       foreach ($CompletedPreviouslyFailedRcdNum in $arrCommonRcdNumsInCompletedandFailed) {
         #Add (and save for logging) Failed but now succeeded Rcd and the last FailureCount
         If ($hshFailedThenSuceededRcds.ContainsKey($CompletedPreviouslyFailedRcdNum)) {
           Write-DMLog -Text ("$(Get-DMScriptName)(): Error: Trying to add duplicate Key: {$CompletedPreviouslyFailedRcdNum} to Hash Table:{hshFailedThenSuceededRcds}")
          } Else { #This Record has not already failed to Add it
           #Increment the FailureCount for this Record
           $hshFailedThenSuceededRcds.Add($RecordArray[$CompletedPreviouslyFailedRcdNum],($hshFailingJobsTracker.$CompletedPreviouslyFailedRcdNum))
          }
         #Remove $FailedJobRcdNum from $hshFailingJobsTracker
         $hshFailingJobsTracker.Remove($CompletedPreviouslyFailedRcdNum)
        }
      }
     $CompletedJobs | Remove-Job

     #################################
     #Remove and track any failed Jobs
     #################################
     $FailedJobs = Get-Job -Name "$MTJobPrefix*" | where {($_.state -eq "Failed")}
     #Use hshFailingJobsTracker to maintain a hash table of failed (via the RcdNum) and their count. RcdNum is the index into $RecordArray
     foreach ($FailedJob in $FailedJobs) {
       #If a job has failed it will return an ErrorRecord PS Object type by redirecting ErrOut to StdOut
       $objFailureResults =  Receive-Job -Job $FailedJob 2>&1
       #Get RcdNum the record that was being processed in the Failed Job
       $FailedJobRcdNum = $FailedJob.RcdNum
       #Capture the ErrorCode, short ErrorMessage and DateTime Stamp
       $objJobFailureErrors = new-object PSObject
       $objJobFailureErrors | add-member -membertype NoteProperty -name "Rcd" -Value $RecordArray[$FailedJobRcdNum]
       $objJobFailureErrors | add-member -membertype NoteProperty -name "JobErrorCode" -Value $objFailureResults.Exception.ErrorCode
       $objJobFailureErrors | add-member -membertype NoteProperty -name "JobErrorMsg" -Value $objFailureResults.Exception.TransportMessage
       $objJobFailureErrors | add-member -membertype NoteProperty -name "JobErrorTimeStamp" -Value $FailedJob.PSEndTime
       $arrObjJobFailureErrors += $objJobFailureErrors
       #If this RcdNum already is being tracked as a failing Job
       If ($hshFailingJobsTracker.ContainsKey($FailedJobRcdNum)) {
         #Increment the FailureCount for this Record
         [Int]$hshFailingJobsTracker.$FailedJobRcdNum += 1
        } Else { #This Record has not already failed to Add it
         $hshFailingJobsTracker.Add($FailedJobRcdNum,1)
        }
       #If, for this record, we've reached the maximum number of job failures allowed
       If ([Int]$hshFailingJobsTracker.$FailedJobRcdNum -ge $MaxJobFailuresPerRecordAllowed) {
         #Add Record to FailedRcd array
         $arrFailedRcds += $RecordArray[$FailedJobRcdNum]
         #Remove $FailedJobRcdNum from $hshFailingJobsTracker
         $hshFailingJobsTracker.Remove($FailedJobRcdNum)
        }
      }
     #Remove the jobs now that we've processed them
     $FailedJobs | Remove-Job

     #################################
     #Remove and track any hung jobs
     #################################
     #Get array of Jobs that have been running for more than $HungJobThresholdinSeconds seconds
     $HungJobs = (Get-Job -Name "$MTJobPrefix*" | where {($_.PSBeginTime -lt ((Get-Date)-(New-TimeSpan -seconds $HungJobThresholdinSeconds)))}) 
     #Add Record to Hung Rcd array
     foreach ($HungJob in $HungJobs) {
       #Get RcdNum the record that was being processed in the Failed Job
       $HungJobRcdNum = $HungJob.RcdNum
       $objHungJob = new-object PSObject
       $objHungJob | add-member -membertype NoteProperty -name "Rcd" -Value $RecordArray[$HungJobRcdNum]
       $objHungJob | add-member -membertype NoteProperty -name "HungJobTimeStamp" -Value $HungJob.PSEndTime
       $arrHungRcds += $objHungJob
      }
     #Remove the jobs now that we've processed them
     $HungJobs | Remove-Job -Force
      
     #Write-Progress
     MTWriteProgress $MTStartTime $MTJobPrefix $MTSessionPrefix $MaxThreads $RcdNum $MaxRecordNum $AllowScriptBlockDebugging

     #Now that we've removed the jobs that were completed, hung or failed too many times,
     # if we're still at our maximum number of threads
     If ((get-job -Name "$MTJobPrefix*").count -ge $MaxThreads) {
       #Sleep until we can have room to launch another thread
       $SleepLoopCtr = 0
       $SleepWarningThreshold = 10
       while ((get-job -Name "$MTJobPrefix*" | where {$_.state -eq "Running"}).count -ge $MaxThreads) {
         #Sleep for 1 second to allow CPU to be used to complete the background jobs
         $SleepLoopCtr += 1
         start-sleep -Seconds 1
         #Update displayed status - Time estimates will have changed
         MTWriteProgress $MTStartTime $MTJobPrefix $MTSessionPrefix $MaxThreads $RcdNum $MaxRecordNum $AllowScriptBlockDebugging
         #Provide an update on sleep status if we sleep for more than $SleepWarningThreshold seconds
         If ($SleepLoopCtr -ge $SleepWarningThreshold) {
           $x=Write-DMLog -Text "We've been sleeping for >= $SleepWarningThreshold seconds, waiting for background jobs to complete"
           $SleepLoopCtr = 0
          }
        }
      } #End If at Max threads
    }

   #If we received the BatchSaveSize argument, check to see if we need to write out a batch of results
   If (($BatchSaveSize -gt 0) -and (($arrobjResults).count -gt $BatchSaveSize)) {
     $BatchSaveCount++
     $ResultsFile = "$ResultsDir\$NameQualifier-Batch-$(($BatchSaveCount*$BatchSaveSize).Tostring()).xml"
     $arrobjResults[0..($BatchSaveSize-1)] | Export-Clixml $ResultsFile -NoClobber:$False
     $arrobjResults = $arrobjResults[$BatchSaveSize..(($arrobjResults).count - 1)]
    } Else {
     #If the objects returned from the ScriptBlock contains a property named "SBUniqueResultsFileName"
     if (-not [string]::IsNullOrEmpty($arrobjResults[0].SBUniqueResultsFileName)) {
       ForEach ($objResults in $arrobjResults) {
         #Save the objects in XML files at the calculated location below
         $ResultsFile = "$ResultsDir\$NameQualifier-$($objResults.SBUniqueResultsFileName).xml"
         $objResults | Export-Clixml $ResultsFile -NoClobber:$False
         $blnSavingResultsToXML = $true
        }
       $arrobjResults = @()
      } #If SBUniqueResultsFileName
    } # If BatchSaveSize
  } #End of Main While Loop
  
 #If we're saving results to result files and we have any leftover results
 If (($BatchSaveSize -gt 0) -and (($arrobjResults).count -gt 0)) {
   $ResultsFile = "$ResultsDir\$NameQualifier-Batch-$((($BatchSaveCount*$BatchSaveSize)+(($arrobjResults).count)).Tostring()).xml"
   $arrobjResults | Export-Clixml $ResultsFile -NoClobber:$False
  }

 #If we're saving one result file per record in passed array
 if ($blnSavingResultsToXML) {
   $arrobjResults = @()
   $x=Write-DMLog -Text "Retrieving Results stored in Rcd specific XML files: {$ResultsDir}"
   $arrobjResults += gci "$ResultsDir\$NameQualifier-*.xml" | Import-Clixml
  }

 #Clean out any leftover Jobs and PSSessions
 Get-Job -Name "$MTJobPrefix*" | Remove-Job
 Get-PSSession -Name "$MTSessionPrefix*" | Remove-PSSession
 
 #Calculate and display Results Metrics
 $MTEndTime = Get-Date
 $ElapsedMins = ($MTEndTime - $MTStartTime).totalminutes
 $ProcessingRate = $MaxRecordNum / $ElapsedMins

 #Remove any Null Record
 $arrobjResults = $arrobjResults | Where-Object {$_}
 $arrFailedRcds = $arrFailedRcds | Where-Object {$_}
 $arrHungRcds = $arrHungRcds | Where-Object {$_}
 $arrObjJobFailureErrors = $arrObjJobFailureErrors | Where-Object {$_}

 #Log Results
 $MTLogFile = Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): ########################################################################" -ReturnLogFile
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): MTEndTime:     {$MTEndTime}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): Completed Invoke-DMMultiThreadingEngine() with the following Results:"
 Write-DMLog -Text ("Invoke-DMMultiThreadingEngine():   Processing Rate {0:N1} (Records/minute)" -f $ProcessingRate)
 Write-DMLog -Text ("Invoke-DMMultiThreadingEngine():   Total Time to complete {0:N2} (Minutes)" -f $ElapsedMins)
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Number of Records to Process  : {$MaxRecordNum}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Number of Records Returned    : {$(($arrobjResults).count)}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Number of Records Failed      : {$(($arrFailedRcds).count)}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Number of Records Hung/Killed : {$(($arrHungRcds).count)}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Number of Job Errors          : {$(($arrObjJobFailureErrors).count)}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine():   Log File : {$MTLogFile}"
 Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): ########################################################################"

 #Create Stats object to return
 $objFcnStats = new-object PSObject
 $objFcnStats | add-member -membertype NoteProperty -name "StartTime" -Value $MTStartTime
 $objFcnStats | add-member -membertype NoteProperty -name "EndTime" -Value $MTEndTime
 $objFcnStats | add-member -membertype NoteProperty -name "ElapsedMins" -Value $ElapsedMins
 $objFcnStats | add-member -membertype NoteProperty -name "LogFile" -Value $MTLogFile
 $objFcnStats | add-member -membertype NoteProperty -name "MaxRecordNum" -Value $MaxRecordNum
 $objFcnStats | add-member -membertype NoteProperty -name "ResultsDir" -Value $ResultsDir
 $objFcnStats | add-member -membertype NoteProperty -name "NameQualifier" -Value $NameQualifier
 $objFcnStats | add-member -membertype NoteProperty -name "MaxThreads" -Value $MaxThreads
 $objFcnStats | add-member -membertype NoteProperty -name "BatchSaveSize" -Value $BatchSaveSize
 $objFcnStats | add-member -membertype NoteProperty -name "RunJobsOnRemoteComputers" -Value $RunJobsOnRemoteComputers
 $objFcnStats | add-member -membertype NoteProperty -name "HungJobThresholdinSeconds" -Value $HungJobThresholdinSeconds
 $objFcnStats | add-member -membertype NoteProperty -name "MaxJobFailuresPerRecordAllowed" -Value $MaxJobFailuresPerRecordAllowed
 $objFcnStats | add-member -membertype NoteProperty -name "RecordsPerMinuteProcessed" -Value $ProcessingRate
 $objFcnStats | add-member -membertype NoteProperty -name "ReturnedErrorCode" -Value $ReturnedErrorCode
 $objFcnStats | add-member -membertype NoteProperty -name "ReturnedErrorMsg" -Value $ReturnedErrorMsg

 #If we had any Job failure error
 If ($arrObjJobFailureErrors.count -gt 0) {
   #Provide report Job failure error
   Write-DMLog -Text "Invoke-DMMultiThreadingEngine(): ### Report of Job Failure Error Messages ###"
   $arrObjJobFailureErrors | Group-Object -Property JobErrorMsg | select -Property Count,@{name="JobErrorMsg";expression={$_.name}}  | sort -Property count -Descending |ft -Property Count,JobErrorMsg -AutoSize -Wrap | out-string | Write-DMLog
  }

 #Close Progress Display
 Write-Progress -Activity Completed -Completed

 #Return $arrobjResults
 $objReturnResult = new-object PSObject
 $objReturnResult | add-member -membertype NoteProperty -name "arrobjResults" -Value $arrobjResults  
 $objReturnResult | add-member -membertype NoteProperty -name "arrFailedRcds" -Value ($arrFailedRcds | sort)
 $objReturnResult | add-member -membertype NoteProperty -name "arrHungRcds" -Value ($arrHungRcds | sort)
 $objReturnResult | add-member -membertype NoteProperty -name "hshFailedThenSuceededRcds" -Value ($hshFailedThenSuceededRcds | Sort)
 $objReturnResult | add-member -membertype NoteProperty -name "arrObjJobFailureErrors" -Value ($arrObjJobFailureErrors)
 $objReturnResult | add-member -membertype NoteProperty -name "objFcnStats" -Value ($objFcnStats)
 Return $objReturnResult
 
} #End Function Invoke-DMMultiThreadingEngine

#Define Functions Aliases as needed
New-Alias -Name Parse-IniFile -Value Parse-DMIniFile -Description "Used to help transition any DMI code currently using Parse-IniFile" -ErrorAction SilentlyContinue
New-Alias -Name Run-DMSQLCmd -Value Get-DMSQLQuery -Description "Used to help transition any DMI code currently using Get-DMSQLQuery" -ErrorAction SilentlyContinue

Export-ModuleMember -Alias *
Export-ModuleMember -Function * 

