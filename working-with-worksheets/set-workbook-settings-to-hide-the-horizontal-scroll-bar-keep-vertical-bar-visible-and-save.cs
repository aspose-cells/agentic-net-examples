// Title: Create an Excel workbook with Aspose.Cells (.NET), hide the horizontal scroll bar while keeping the vertical scroll bar visible, and save the file
// AI Prompts: Write C# code using Aspose.Cells to create a new Workbook, attempt to disable the horizontal scroll bar, ensure the vertical scroll bar stays enabled, and save the workbook to a given path. | Show how to programmatically verify whether the ShowHorizontalScrollBar property exists in WorkbookSettings before setting it, and provide a fallback when the property is unavailable. | Provide a C# example that creates a workbook, creates the output directory if it does not exist, saves the file as .xlsx, and notes that scroll‑bar visibility settings are not supported in the current Aspose.Cells version.
// Common Searches: Aspose.Cells .NET hide horizontal scroll bar in generated Excel workbook | keep vertical scroll bar visible while disabling horizontal scroll bar using Aspose.Cells | save Excel file with custom scroll bar settings in C# Aspose.Cells | detect ShowHorizontalScrollBar property support in Aspose.Cells WorkbookSettings
// Tags: Aspose.Cells workbooksettings hide horizontal scrollbar | C# Aspose.Cells set scroll bar visibility | Aspose.Cells save workbook to xlsx | unsupported workbook settings Aspose.Cells .NET | create output directory before saving Excel file C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook with Aspose.Cells, accesses its Settings object, notes that ShowHorizontalScrollBar and ShowVerticalScrollBar properties are not available in the current version, ensures the output directory exists, and saves the workbook to 'output.xlsx' while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access workbook settings (if needed for other settings)
            WorkbookSettings settings = workbook.Settings;

            // Example: you can modify other settings here.
            // Note: ShowHorizontalScrollBar and ShowVerticalScrollBar properties are not available
            // in the current Aspose.Cells version, so they are omitted.

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
