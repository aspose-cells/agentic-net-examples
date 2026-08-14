// Title: Save a Modified Aspose.Cells Workbook to a New XLSX File While Preserving Formatting
// Description: Demonstrates how to modify cells in a new Aspose.Cells Workbook (C#) and save it as a separate XLSX file using SaveFormat.Xlsx, ensuring that all original styles, textures, and formatting are retained.
// Keywords: Aspose.Cells save workbook C# | preserve formatting Aspose.Cells | Workbook.Save new file | SaveFormat.Xlsx example | .NET Excel export preserving styles | modify cells and keep layout
// Common Searches: Aspose.Cells save workbook without losing formatting | C# Aspose.Cells save as new XLSX file | how to keep cell styles when saving Excel with Aspose | Workbook.Save preserving textures Aspose.Cells | duplicate workbook and save copy .NET
// Developer Intent: The developer needs to write a C# program that updates specific cells in a workbook and writes the result to a new XLSX file without altering any existing formatting or visual elements.
// Use Cases: Create a template workbook, inject fresh data, and export a styled copy for distribution. | Automate daily reporting by updating values in a master file and saving a versioned report that retains the original design. | Implement a workflow that clones a workbook, applies business‑logic changes, and stores the modified version separately while preserving all cell formatting.
// AI Prompts: Generate C# code that opens an Aspose.Cells workbook, changes cells A1, B2, C3, and saves the result as a new XLSX file while keeping all formatting intact. | Explain the role of the SaveFormat enumeration in Aspose.Cells and how to use Workbook.Save to retain styles and textures. | Show how to duplicate a worksheet with its full formatting, modify data, and export the duplicate as an independent workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    // Demonstrates how to modify cells in a new Aspose.Cells Workbook (C#) and save it as a separate XLSX file using SaveFormat.Xlsx, ensuring that all original styles, textures, and formatting are retained.
    public class SaveModifiedWorkbook
    {
        public static void Run()
        {
            try
            {
                // Initialize a new workbook (default format is Xlsx)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Example modification: change some cell values
                sheet.Cells["A1"].PutValue("Original");
                sheet.Cells["B2"].PutValue(12345);
                sheet.Cells["C3"].PutValue(DateTime.Now);

                // Save with the same format preserving formatting
                string outputPath = "ModifiedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SaveModifiedWorkbook.Run();
        }
    }
}
