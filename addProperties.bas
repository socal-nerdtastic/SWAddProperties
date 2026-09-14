' exactly the same as the .swp, just in github-readable format

Option Explicit

Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim sFilePath As String
Dim sFileName As String

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    sFileName = Mid(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") + 1)
    sFilePath = swModel.GetPathName

'Properties to delete
    'swModel.DeleteCustomInfo "Description"
    'swModel.DeleteCustomInfo "Number"

'Properties to add
    swModel.AddCustomInfo2 "Number", swCustomInfoText, ""
    swModel.AddCustomInfo2 "Description", swCustomInfoText, ""
    swModel.AddCustomInfo2 "Material", swCustomInfoText, "SEE NOTES"
    swModel.AddCustomInfo2 "Finish", swCustomInfoText, "SEE NOTES"
    swModel.AddCustomInfo2 "Manufacturer", swCustomInfoText, ""
    swModel.AddCustomInfo2 "Manufacturer P/N", swCustomInfoText, ""
    swModel.AddCustomInfo2 "Revision", swCustomInfoText, "1"
    swModel.AddCustomInfo2 "Bin", swCustomInfoText, ""
    swModel.AddCustomInfo2 "Unit Of Measure", swCustomInfoText, "EA"
    'swModel.AddCustomInfo2 "Weight", swCustomInfoText, Chr(34) + "SW-Mass" + "@" + sFileName + Chr(34)
    'swModel.AddCustomInfo2 "Location", swCustomInfoText, "$" + "PRP:" + Chr(34) + Chr(83) + Chr(87) + "-Folder Name" + Chr(34)

'Optional: Open the Summary Information window
    swModel.FileSummaryInfo

End Sub
