
<#
.Synopsis
   Helper class for C# code.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

$csharpSource = @"
public static class HashCodeUtility {

    public static int GetDeterministicHashCode(this string str)
    {
        unchecked
        {
            int hash1 = (5381 << 16) + 5381;
            int hash2 = hash1;

            for (int i = 0; i < str.Length; i += 2)
            {
                hash1 = ((hash1 << 5) + hash1) ^ str[i];
                if (i == str.Length - 1)
                    break;
                hash2 = ((hash2 << 5) + hash2) ^ str[i + 1];
            }

            return hash1 + (hash2 * 1566083941);
        }
    }

    public static int UAdd(int i1, int i2) {
        unchecked 
        {
            return i1+i2;
        }
    }
}
"@

Add-Type -TypeDefinition $csharpSource
