// Title: Create a macro‑enabled .xlsm workbook and embed a Workbook_Open VBA macro to format header cells with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to create an .xlsm file, add sample data, enable macros, and inject a Workbook_Open VBA procedure that makes the header row bold, applies a light blue fill, and adjusts column widths automatically. | Show how to retrieve or add the ThisWorkbook VBA module in an Aspose.Cells workbook and assign VBA source code to it before saving as a macro‑enabled workbook. | Demonstrate enabling macros, inserting a VBA macro, and saving the workbook in Xlsm format with Aspose.Cells for .NET, including handling the case where the VBA project is not created.
// Common Searches: aspnet create xlsm file with embedded VBA using Aspose.Cells | how to add Workbook_Open macro programmatically in C# Aspose.Cells | enable macros and save as Xlsm format with Aspose.Cells .NET example | add ThisWorkbook module and set VBA code in Aspose.Cells C# | auto‑fit columns and format header row via VBA macro in Aspose.Cells workbook
// Tags: Aspose.Cells macro-enabled workbook creation | C# add VBA module Aspose.Cells | Workbook_Open event VBA injection Aspose.Cells | save workbook as Xlsm Aspose.Cells .NET | format header row via VBA macro Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example creates a new workbook, enables macro support, adds sample data to a sheet named "Data", injects a Workbook_Open VBA routine that bolds the header row, applies a light blue background, and adjusts column widths, then saves the file as a macro‑enabled .xlsm workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and enable macro support
            Workbook workbook = new Workbook();
            workbook.Settings.EnableMacros = true;

            // Ensure a VBA project exists (created automatically when macros are enabled)
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
                throw new InvalidOperationException("VBA project could not be created.");

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // VBA code that runs when the workbook is opened
            string vbaCode = @"
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(""Data"")
    ws.Range(""A1:B1"").Font.Bold = True
    ws.Range(""A1:B1"").Interior.Color = RGB(200, 200, 255)
    ws.Columns(""A:B"").AutoFit
End Sub
";

            // Get (or add) the ThisWorkbook module and set its code
            VbaModule thisWorkbookModule = null;
            try
            {
                // Try to retrieve the module by name (indexer returns VbaModule)
                thisWorkbookModule = vbaProject.Modules["ThisWorkbook"];
            }
            catch
            {
                // Ignored – will add the module if not found
            }

            if (thisWorkbookModule == null)
            {
                // Add returns the index of the new module
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Document, "ThisWorkbook");
                thisWorkbookModule = vbaProject.Modules[moduleIndex];
            }

            thisWorkbookModule.Codes = vbaCode;

            // Determine output path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroEnabledWorkbook.xlsm");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as a macro‑enabled file
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
