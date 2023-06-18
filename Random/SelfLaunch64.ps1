#Check 64 Bit Script
#$env:PROCESSOR_ARCHITEW6432 Returns AMD64 on 32 bit Powershell and blank on 64 bit if system is 64bit
if($env:PROCESSOR_ARCHITEW6432-eq "AMD64") {
    #Run PS command from 64 bit powershell. on 32 bit powershell sysnative returns path
    &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
    #Exit after re-running full script in 64 bit subset
    Exit
}
