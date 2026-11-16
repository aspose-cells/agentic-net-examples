// Title: How to apply a LightBlue background to a header row and freeze the first row in an Excel file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set background color for first row and keep it visible while scrolling | C# freeze panes after styling header row with solid fill using Aspose.Cells | Apply solid fill style to A1:F1 and freeze top row in Excel with Aspose.Cells .NET | How to create a styled header and freeze top row in an Aspose.Cells workbook | Set header background and freeze first row programmatically with Aspose.Cells for .NET
// Tags: header row background color using Aspose.Cells API | freeze panes first row Aspose.Cells .NET | solid fill style application Aspose.Cells range | styled header creation Aspose.Cells workbook | worksheet freeze panes after styling header

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, defines a LightBlue solid fill style for cells A1:F1, applies the style to the header range, freezes the first row with FreezePanes, ensures the output directory exists, and saves the file as HeaderWithFreeze.xlsx while handling potential exceptions.
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

            // Define the header range (first row, columns A to F)
            const string headerRange = "A1:F1";

            // Create a style for the header background
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;      // Desired background color
            headerStyle.Pattern = BackgroundType.Solid;        // Apply solid fill

            // Apply the style to the header range
            Aspose.Cells.Range range = sheet.Cells.CreateRange(headerRange);
            range.ApplyStyle(headerStyle, new StyleFlag { All = true });

            // Freeze the first row (row 1) while keeping columns unfrozen
            // Parameters: row (first unfrozen row), column (first unfrozen column), totalRows to freeze, totalColumns to freeze
            sheet.FreezePanes(1, 0, 1, 0);

            // Define output file path
            string outputPath = "HeaderWithFreeze.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
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
