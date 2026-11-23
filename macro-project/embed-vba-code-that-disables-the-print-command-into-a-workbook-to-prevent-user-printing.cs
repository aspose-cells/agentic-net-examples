// Title: How to embed a VBA Workbook_BeforePrint macro that blocks printing in an XLSM file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add VBA Workbook_BeforePrint event to stop printing | programmatically disable Excel print command in a macro-enabled workbook with Aspose.Cells | inject VBA code to cancel print operation in an XLSM using Aspose.Cells .NET | save workbook as macro-enabled file and show message when user tries to print | Aspose.Cells create XLSM with VBA that prevents user printing
// Tags: Aspose.Cells VBA module injection C# | Workbook_BeforePrint event cancellation | create macro-enabled XLSM with Aspose.Cells | disable Excel printing programmatically .NET | add VBA Document module Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace Example
{
    // The example creates a new Workbook, accesses its VbaProject, adds a Document module named "ThisWorkbook", injects a Workbook_BeforePrint VBA routine that sets Cancel = True and displays a message, and then saves the workbook as a macro‑enabled XLSM file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Add a Document module named "ThisWorkbook"
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Document, "ThisWorkbook");
                VbaModule thisWorkbookModule = vbaProject.Modules[moduleIndex];

                // VBA code that cancels any print operation and shows a message
                thisWorkbookModule.Codes = @"
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Cancel = True
    MsgBox ""Printing is disabled.""
End Sub
";

                // Save the workbook as a macro‑enabled file (XLSM)
                string outputPath = "PrintDisabled.xlsm";
                workbook.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
