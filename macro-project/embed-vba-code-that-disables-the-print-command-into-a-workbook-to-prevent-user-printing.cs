// Title: Embed VBA to Block Printing in an Excel Workbook with Aspose.Cells for .NET
// Description: This example shows how to create a macro‑enabled workbook, inject a VBA Workbook_BeforePrint routine that sets Cancel = True, optionally protect the VBA project with a password, and save the file as an .xlsm using Aspose.Cells for .NET.
// Keywords: Aspose.Cells embed VBA | disable Excel printing programmatically | Workbook_BeforePrint Aspose .NET | create macro enabled workbook C# | protect VBA project Aspose.Cells
// Common Searches: add VBA code to prevent printing with Aspose.Cells | how to embed Workbook_BeforePrint event in .xlsm using C# | save macro‑enabled Excel file after inserting VBA via Aspose | protect VBA project when adding code programmatically
// Developer Intent: Insert a VBA routine that cancels every print request and store the workbook as a macro‑enabled file.
// Use Cases: Distribute confidential reports that cannot be printed. | Provide Excel templates that enforce a no‑print policy while allowing other macros. | Lock the VBA project after adding anti‑print code to reduce tampering.
// AI Prompts: Write C# code with Aspose.Cells that adds a Workbook_BeforePrint handler to cancel printing and then protects the VBA project. | Show the steps to create an .xlsm file, embed anti‑print VBA, and save it using Aspose.Cells for .NET. | Explain how to test that the embedded VBA disables the Print command when the workbook is opened.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

// This example shows how to create a macro‑enabled workbook, inject a VBA Workbook_BeforePrint routine that sets Cancel = True, optionally protect the VBA project with a password, and save the file as an .xlsm using Aspose.Cells for .NET.
class EmbedVbaDisablePrint
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Ensure a VBA project exists by saving as a macro‑enabled file and reloading it
        string tempPath = "temp.xlsm";
        wb.Save(tempPath, SaveFormat.Xlsm);
        wb = new Workbook(tempPath);
        System.IO.File.Delete(tempPath);

        // The first module in the VBA project is the ThisWorkbook module
        VbaModule thisWorkbook = wb.VbaProject.Modules[0];

        // VBA code that cancels any print operation
        StringBuilder vbaCode = new StringBuilder();
        vbaCode.AppendLine("Private Sub Workbook_BeforePrint(Cancel As Boolean)");
        vbaCode.AppendLine("    Cancel = True");
        vbaCode.AppendLine("End Sub");

        // Assign the code to the ThisWorkbook module
        thisWorkbook.Codes = vbaCode.ToString();

        // Protect the VBA project (optional, not locked for viewing)
        wb.VbaProject.Protect(false, "vbaPassword");

        // Save the workbook as a macro‑enabled file
        wb.Save("Workbook_NoPrint.xlsm", SaveFormat.Xlsm);
    }
}
