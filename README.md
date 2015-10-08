# VBS_Tools
Function specific utility tools in vbscript

set_assemblyversion
  Takes up to two command line parameters
  The first one specifies the path to the file AssemblyInfo.cs
  e.g. - ./Project/Properties/AssemblyInfo.cs
  The second parameter is the number to use in the 3rd place of the AssemblyFileVersion field - usually the build number.
  If the environment variable BUILD_NUMBER is set, then this value will be used to set the 3rd field of the version value in AssemblyFileVersion instead of the 2nd command line parameter.

15mins
  Duecedly simple script to output 24 hours in 15 minute increments.
  Useful for populating a column in excel, just copy and paste.


