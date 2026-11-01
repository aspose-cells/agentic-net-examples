// Title: How to apply a custom number format (0.0%) to display percentages with one decimal place using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# format cell as percentage with one decimal place | set custom number format 0.0% in Excel workbook using Aspose.Cells .NET | apply percentage style with one decimal to a cell in Aspose.Cells | C# Aspose.Cells custom number format for percentages | display decimal value as 12.3% in Excel using Aspose.Cells
// Tags: percentage custom format 0.0% Aspose.Cells | percentage formatting one decimal C# | apply style to cell Aspose.Cells | Excel workbook custom format .NET | Aspose.Cells percentage style example

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes a decimal to cell A1, defines a style with the custom number format "0.0%" to show percentages with one decimal place, applies the style to the cell, ensures the output directory exists, and saves the file as PercentageFormatted.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a decimal value in cell A1 (e.g., 0.1234 will display as 12.3%)
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(0.1234);

            // Create a style and set a custom number format for percentages with one decimal place
            Style style = workbook.CreateStyle();
            style.Custom = "0.0%";

            // Apply the style to the cell (entire style is safe here)
            cell.SetStyle(style);

            // Define output file path
            string outputPath = "PercentageFormatted.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
