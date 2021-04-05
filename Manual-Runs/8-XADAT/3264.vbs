call CSI_GetOSBits() 

wscript.echo CSI_GetOSBits

Function CSI_GetOSBits() 

CSI_GetOSBits = GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth 

  End Function

