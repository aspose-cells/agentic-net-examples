// Title: Create a macro‑enabled worksheet and save it as an .xlsm file using Aspose.Cells for .NET (C#)
// Common Searches: C# code to generate a macro‑enabled .xlsm workbook with Aspose.Cells | how to rename the default sheet before saving as .xlsm using Aspose.Cells | Aspose.Cells save workbook as macro‑enabled file programmatically | example of creating an .xlsm file in .NET with Aspose.Cells | Aspose.Cells create macro‑enabled workbook without using Excel Interop
// Tags: create macro-enabled workbook Aspose.Cells C# | save workbook as xlsm Aspose.Cells | rename worksheet Aspose.Cells | macro-enabled file generation .NET Aspose.Cells | Aspose.Cells workbook creation without Interop

using System;
using System.IO;
using Aspose.Cells;

// The sample demonstrates how to instantiate a new Workbook, rename its default worksheet, and save it as a macro‑enabled .xlsm file using Aspose.Cells for .NET in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (will be saved as macro‑enabled)
            Workbook workbook = new Workbook();

            // Rename the default worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Save the workbook as a macro‑enabled file
            string outputPath = "MacroEnabledWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
