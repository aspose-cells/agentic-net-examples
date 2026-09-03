// Title: Freeze the first column of a worksheet and export the workbook as a macro‑enabled XLSM file with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to freeze column A on a worksheet and then saves the workbook in macro‑enabled XLSM format. | Generate a .NET example that applies FreezePanes to the header column and writes the file using SaveFormat.Xlsm.
// Common Searches: C# Aspose.Cells freeze first column while scrolling | how to save a workbook as macro enabled XLSM using Aspose.Cells .NET | freeze panes header columns and export as XLSM Aspose.Cells example | Aspose.Cells SaveFormat.Xlsm usage in C#
// Tags: freeze panes first column Aspose.Cells | export workbook to XLSM Aspose.Cells | macro enabled Excel file generation .NET | worksheet FreezePanes API | SaveFormat.Xlsm example C#

using Aspose.Cells;
using System;
using System.IO;

// // Creates a new workbook, freezes column A on the first worksheet, ensures the output directory exists, and saves the file as a macro‑enabled XLSM using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first column (A) so header columns stay visible while scrolling
            // Parameters: totalRows, totalColumns, rows to freeze, columns to freeze
            sheet.FreezePanes(0, 0, 0, 1);

            // Define output file path
            string outputPath = "Output.xlsm";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook as an XLSM (macro‑enabled) file
            workbook.Save(outputPath, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
