// Title: Hide an Excel worksheet with Aspose.Cells for .NET while keeping it accessible to VBA macros
// AI Prompts: Write C# code that uses Aspose.Cells to set a worksheet's Visibility to VeryHidden so the sheet is hidden from the UI but can still be referenced by VBA. | Show a complete example that creates a workbook, renames the first sheet, hides it, writes data to a cell on the hidden sheet, ensures the output folder exists, and saves the file as .xlsx using Aspose.Cells. | Demonstrate how to switch a worksheet between Visible, Hidden, and VeryHidden states in C# with Aspose.Cells and explain the impact on VBA macro access.
// Common Searches: asp.net hide worksheet veryhidden using Aspose.Cells C# | keep hidden Excel sheet accessible to VBA macros Aspose.Cells | Aspose.Cells set worksheet IsVisible false vs VeryHidden | C# create workbook and hide first sheet before saving with Aspose.Cells | save Excel file with hidden sheet Aspose.Cells .NET
// Tags: Aspose.Cells set worksheet VeryHidden | C# hide Excel sheet Aspose.Cells | Aspose.Cells workbook save xlsx hidden sheet | VBA access very hidden worksheet Aspose.Cells | ensure output directory exists C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a workbook, renames the first worksheet, marks it as VeryHidden (or sets IsVisible to false), adds a value to a hidden cell, guarantees the output directory exists, and saves the file as HiddenWorksheet.xlsx, allowing VBA macros to still interact with the hidden sheet.
class HideWorksheetExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a default workbook with one sheet

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Hide the worksheet from the UI.
            // For newer Aspose.Cells versions you can use sheet.Visibility = SheetVisibility.VeryHidden;
            // Here we use IsVisible to ensure compatibility with all versions.
            sheet.IsVisible = false;

            // Optionally, add some data to the hidden sheet
            sheet.Cells["A1"].PutValue("Secret Data");

            // Define output file path
            string outputPath = "HiddenWorksheet.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
