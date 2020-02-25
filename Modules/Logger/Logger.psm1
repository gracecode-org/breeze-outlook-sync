<#
.Synopsis
   Class Module for logging
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

class Logger {

    static [Logger] $Logger = [Logger]::new()

    static [int] $LOGLEVEL_QUIET=0
    static [int] $LOGLEVEL_NORMAL=1
    static [int] $LOGLEVEL_DEBUG=2

    [string] $LogPath = $env:APPDATA + "\BreezeOutlookSync\logs"
    [string] $LogFileName = "sync"
    [string] $LogFileExt = ".log"
    [string] $LogFile =  $env:APPDATA + "\BreezeOutlookSync\logs\sync.log"
    [int] $LogLevel = [Logger]::LOGLEVEL_NORMAL
    [int] $MaxLogFiles = 20
    [int] $MaxLogSize = 20971520 #20mb
    
    static [void] init([string] $logPath, [int] $logLevel, [int] $maxLogSize, [int] $maxLogFiles) {
        [Logger]::Logger.LogPath = $logPath
        [Logger]::Logger.LogFile = $logPath + "\sync.log"
        [Logger]::Logger.LogLevel = $logLevel
        [Logger]::Logger.MaxLogSize = $maxLogSize
        [Logger]::Logger.MaxLogFiles = $maxLogFiles
        [Logger]::Logger.CreateLogsDir()
        [Logger]::Logger.Rotate()
    }

    hidden [void] Rotate() {
        $currentSize = (Get-Item ($this.LogFile)).Length
        # if MaxLogFiles is 1 just keep the original one and let it grow
        if (-not($this.MaxLogFiles -eq 1)) {
            if ($currentSize -ge $this.MaxLogSize) {

                $newLogFileName = $this.LogFileName + (Get-Date -Format 'yyyyMMddHHmmss').ToString() + $this.LogFileExt

                Copy-Item -Path $this.LogFile -Destination (Join-Path (Split-Path $this.LogFile) $newLogFileName)

                Clear-Content $this.LogFile

                # if MaxLogFiles is 0 don't delete any old archived log files
                if (-not($this.MaxLogFiles -eq 0)) {

                    # set filter to search for archived log files
                    $archivedLogFileFilter = $this.LogFileName + '??????????????' + $this.LogFileExt

                    # get archived log files
                    $oldLogFiles = Get-Item -Path "$(Join-Path -Path $this.LogPath -ChildPath $archivedLogFileFilter)"

                    if ([bool]$oldLogFiles) {
                        # compare found log files to MaxLogFiles parameter of the log object, and delete oldest until we are
                        # back to the correct number
                        if (($oldLogFiles.Count + 1) -gt $this.MaxLogFiles) {
                            [int]$numTooMany = (($oldLogFiles.Count) + 1) - $this.MaxLogFiles
                            $oldLogFiles | Sort-Object 'LastWriteTime' | Select-Object -First $numTooMany | Remove-Item
                        }
                    }
                }
            }
        }
    }
    
    hidden CreateLogsDir() {
        if (-not (Test-Path -PathType Container ($this.LogPath))) {
            New-Item -ItemType Directory -Path ($this.LogPath)
        }
    }

    static [void] Write([string] $logstring) {
        $datetime = Get-Date -Format "MM/dd/yyyy HH:mm:ss.fff"
        Add-content ([Logger]::Logger.LogFile) -value  "$datetime $logstring"
    }

    static [void] Write([string] $logstring, [boolean] $toConsole) {
        if($toConsole) {
            Write-Host $logstring
        }
        [Logger]::Write($logstring)
    }
    
    static [void] Write([string] $logstring, [boolean] $toConsole, [int] $level) {
        if($level -le ([Logger]::Logger.LogLevel)) {
            if($toConsole) {
                Write-Host $logstring
            }
            [Logger]::Write($logstring)
        }
    }
}