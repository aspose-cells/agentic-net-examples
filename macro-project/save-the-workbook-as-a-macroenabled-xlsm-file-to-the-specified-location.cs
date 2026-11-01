// Title: Saving a new workbook as a macro‑enabled XLSM file to a custom folder with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to save workbook as macro‑enabled XLSM in a specific directory | How to use SaveFormat.Xlsm with Aspose.Cells to export a macro‑enabled Excel file | Create output folder if it does not exist before saving XLSM using Aspose.Cells C# | Write values to cells and generate a macro‑enabled workbook with Aspose.Cells .NET | Aspose.Cells save workbook as .xlsm programmatically in C#
// Tags: save workbook as XLSM using Aspose.Cells | Aspose.Cells SaveFormat.Xlsm usage | create macro‑enabled Excel file C# | ensure output directory exists C# | write cell values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes sample data to cells A1 and B1, checks that the target folder exists (creating it if necessary), and saves the file as a macro‑enabled XLSM using Aspose.Cells' SaveFormat.Xlsm.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(12345);

            // Define output folder and file path
            string folderPath = @"C:\Path\To\Your\Folder";
            string outputPath = Path.Combine(folderPath, "MacroEnabledWorkbook.xlsm");

            // Ensure the output directory exists
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
