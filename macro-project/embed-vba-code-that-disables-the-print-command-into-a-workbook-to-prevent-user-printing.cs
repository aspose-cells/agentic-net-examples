// Title: Inject a VBA Workbook_BeforePrint macro to block printing in an Excel XLSM file using Aspose.Cells for .NET (C#)
// Common Searches: how to programmatically add a Workbook_BeforePrint VBA handler with Aspose.Cells C# | C# code to embed a macro that cancels printing in an Excel XLSM using Aspose.Cells | save an Excel workbook with a disabled print command via Aspose.Cells VBA project | prevent users from printing an Excel file by injecting VBA through Aspose.Cells .NET
// Tags: inject VBA Workbook_BeforePrint macro using Aspose.Cells C# | disable Excel printing via macro-enabled XLSM generation | add VBA module to workbook with Aspose.Cells .NET | save macro-enabled XLSM file programmatically Aspose.Cells | prevent user printing in Excel through VBA injection

using Aspose.Cells;
using Aspose.Cells.Vba;

// Create a new workbook
Workbook workbook = new Workbook();

// Access the VBA project of the workbook
VbaProject vbaProject = workbook.VbaProject;

// Get the ThisWorkbook module (code-behind for the workbook)
VbaModule thisWorkbookModule = vbaProject.Modules["ThisWorkbook"];

// Insert VBA code that cancels any print operation
thisWorkbookModule.Codes = 
    "Private Sub Workbook_BeforePrint(Cancel As Boolean)\n" +
    "    Cancel = True\n" +
    "    MsgBox \"Printing is disabled.\"\n" +
    "End Sub";

// Save the workbook as a macro‑enabled file (XLSM)
XlsSaveOptions saveOptions = new XlsSaveOptions(SaveFormat.Xlsm);
workbook.Save("PrintDisabledWorkbook.xlsm", saveOptions);
