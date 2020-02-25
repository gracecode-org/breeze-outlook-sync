<#
.Synopsis
   Class Module representing configuration data.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>


class Config {

    [PSCustomObject] $cfg;

    Config([string] $configJsonFile) {
        $this.cfg = Get-Content ($configJsonFile) | Out-String | ConvertFrom-Json
    }

    [PSCustomObject] GetConfigObject() {
        return $this.cfg
    }

    static [string] ReplaceEnvVars([string] $s) {
        $s = $s.replace('%APPDATA%', $env:APPDATA)
        $s = $s.replace('%LOCALAPPDATA%', $env:LOCALAPPDATA)
        return $s
    }

    [string] GetCachePath() {
        if([string]::IsNullOrEmpty($this.cfg.CachePath)) {
            $this.cfg.CachePath = "%APPDIR%\BreezeOutlookSync\cache"
        }
        return [Config]::ReplaceEnvVars($this.cfg.CachePath)
    }

    [string] GetLogPath() {
        if([string]::IsNullOrEmpty($this.cfg.LogPath)) {
            $this.cfg.LogPath = "%APPDIR%\BreezeOutlookSync\logs"
        }
        return [Config]::ReplaceEnvVars($this.cfg.LogPath)
    }

}