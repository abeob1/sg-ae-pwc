'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On


Namespace AddOnInstaller
    
    Public Class AddOnInstallInfo
        
        Public StrAddOnInstallPath As String
        
        Public AddOnName As String
        
        Public PartnerName As String
        
        Public ExeFile As String
        
        Public DIFile As String
        
        Public UIFile As String
        
        Public B1WizardBaseFile As String
        
        Public ExtraFiles As System.String()
        
        Public ExtraDirectories As System.String()
        
        Public RestartNeeded As Boolean
        
        Public Sub New()
            MyBase.New
            'INITIALIZATION CODE	...
            AddOnName = "AE_PWC_IN02"
            PartnerName = "Abeo"
            RestartNeeded = False
            ExeFile = "AE_PWC_AO02.exe"
            UIFile = "SAPbouiCOM.dll"
            ExtraFiles = New String() {"AE_PWC_AO02.exe.config", "CFL_Costcenter.srf", "CFL_Distribution.srf", "CostAllocation.srf", "GenerateTextFile.srf", "JournalEntryInfo.rpt", "Logo.bmp", "Summaryreport.rpt"}
            ExtraDirectories = New String() {"", "", "", "", "", "", "", ""}
        End Sub
    End Class
End Namespace
