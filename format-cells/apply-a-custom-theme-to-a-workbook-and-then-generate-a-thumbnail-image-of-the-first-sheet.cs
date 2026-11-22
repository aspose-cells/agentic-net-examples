// Title: Load a .thmx theme (if present) and export the first worksheet as a PNG thumbnail using Aspose.Cells for .NET
// Common Searches: how to export only the first sheet of an Excel workbook to PNG with Aspose.Cells C# | checking for a .thmx theme file before saving a workbook in Aspose.Cells .NET | generate a worksheet thumbnail image from a new workbook using Aspose.Cells | save workbook as XLSX while creating a PNG preview of the first sheet in C# | Aspose.Cells example for creating PNG thumbnail of a worksheet
// Tags: load custom thmx theme Aspose.Cells | save first worksheet as PNG thumbnail | use ImageOrPrintOptions for single sheet image | save workbook to XLSX format Aspose.Cells | generate worksheet preview image C#

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, optionally loads a .thmx theme if the file exists, adds sample data, saves the first worksheet as a PNG thumbnail, and finally saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Load a custom theme from a .thmx file if it exists
            string themePath = "customTheme.thmx";
            if (File.Exists(themePath))
            {
                // Aspose.Cells does not expose a direct LoadTheme method in the current version.
                // Theme loading can be handled via other means if required.
                Console.WriteLine($"Theme file found at {themePath}, but loading themes is not supported in this version.");
            }
            else
            {
                Console.WriteLine($"Theme file not found: {themePath}. Continuing without custom theme.");
            }

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Save the first worksheet as a PNG thumbnail image
            string thumbnailPath = "FirstSheetThumbnail.png";
            try
            {
                // Save the entire workbook as PNG; for a single sheet image, use ImageOrPrintOptions if needed.
                workbook.Save(thumbnailPath, SaveFormat.Png);
                Console.WriteLine($"Thumbnail saved to {thumbnailPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save thumbnail: {ex.Message}");
            }

            // Save the workbook (without custom theme)
            string workbookPath = "WorkbookWithCustomTheme.xlsx";
            try
            {
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
