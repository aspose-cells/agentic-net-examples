// Title: Hide a worksheet from the Excel UI but keep it accessible to VBA macros using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide worksheet but allow VBA macro access | Create macro-enabled XLSM workbook with a hidden sheet using Aspose.Cells | Add VBA module to an Aspose.Cells workbook that reads a hidden worksheet | Set worksheet IsVisible property to false with Aspose.Cells .NET | Generate XLSM file containing hidden data sheet and VBA code via Aspose.Cells
// Tags: worksheet hidden state Aspose.Cells | macro-enabled XLSM creation Aspose.Cells | add VBA module to workbook Aspose.Cells | control worksheet visibility Aspose.Cells | access hidden sheet via VBA Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace HideWorksheetWithVba
{
    // The example demonstrates how to create a new workbook, add a worksheet named "HiddenData", populate it with data, hide the sheet by setting its IsVisible property to false, inject a VBA class module that reads the hidden sheet, and save the file as a macro‑enabled XLSM using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook (default format is Xlsx)
                Workbook workbook = new Workbook();

                // 2. Add a new worksheet that we want to hide
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenData");

                // Populate the hidden sheet with some sample data
                hiddenSheet.Cells["A1"].PutValue("Secret ID");
                hiddenSheet.Cells["B1"].PutValue("Secret Value");
                hiddenSheet.Cells["A2"].PutValue(101);
                hiddenSheet.Cells["B2"].PutValue("Alpha");
                hiddenSheet.Cells["A3"].PutValue(102);
                hiddenSheet.Cells["B3"].PutValue("Beta");

                // 3. Hide the worksheet from the UI (user cannot see it in Excel)
                hiddenSheet.IsVisible = false; // false = hidden, true = visible

                // 4. Add a VBA module that can read data from the hidden sheet
                //    The macro will be stored in the workbook and can be executed
                //    even though the sheet is hidden.
                VbaProject vbaProject = workbook.VbaProject;

                // Add a standard module named "ReadHidden"
                // Use VbaModuleType.Class as the available enum value
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "ReadHidden");
                VbaModule vbaModule = vbaProject.Modules[moduleIndex];

                // VBA code that reads the hidden sheet and shows a message box
                string vbaCode = @"
Sub ShowHiddenData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""HiddenData"")
    Dim msg As String
    msg = ""Data in Hidden Sheet:"" & vbCrLf
    Dim i As Long
    For i = 2 To ws.UsedRange.Rows.Count
        msg = msg & ws.Cells(i, 1).Value & "" : "" & ws.Cells(i, 2).Value & vbCrLf
    Next i
    MsgBox msg, vbInformation, ""Hidden Sheet Content""
End Sub
";
                vbaModule.Codes = vbaCode;

                // 5. Save the workbook as a macro‑enabled file (XLSM)
                string outputPath = "HiddenWorksheetDemo.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("The worksheet 'HiddenData' is hidden from the UI but can be accessed via the VBA macro 'ShowHiddenData'.");
                Console.WriteLine("Open the file in Excel and run the macro to verify.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
