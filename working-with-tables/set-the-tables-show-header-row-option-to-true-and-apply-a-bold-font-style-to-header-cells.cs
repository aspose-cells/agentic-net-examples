// Title: Show Excel table header row and apply bold font to header cells using Aspose.Cells for .NET (C#)
// Common Searches: how to enable ShowHeaderRow for a ListObject and make the header bold with Aspose.Cells C# | Aspose.Cells C# example to display table header row and set bold font on header cells | C# code to set Excel table header visibility and apply bold style using Aspose.Cells | formatting Excel table header row as bold with Aspose.Cells for .NET
// Tags: Aspose.Cells set ListObject ShowHeaderRow true | Aspose.Cells apply bold style to table header | C# format Excel table header bold Aspose.Cells | Aspose.Cells ListObject header row formatting | Excel table header styling Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, checks for a table on the first worksheet, enables the table's ShowHeaderRow property, creates a bold font style, applies the style only to the header row cells of the table, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (worksheet.ListObjects.Count == 0)
                throw new InvalidOperationException("No tables (ListObjects) found in the worksheet.");

            var table = worksheet.ListObjects[0];

            // Ensure the header row is displayed
            table.ShowHeaderRow = true;

            // Create a style with bold font
            var boldStyle = workbook.CreateStyle();
            boldStyle.Font.IsBold = true;

            // Specify that only the font bold attribute should be applied
            var styleFlag = new StyleFlag { FontBold = true };

            // Build a range that represents the header row of the table
            int headerRow = table.StartRow;                     // Header row index
            int startColumn = table.StartColumn;               // First column of the table
            int columnCount = table.EndColumn - table.StartColumn + 1; // Number of columns in the table
            var headerRange = worksheet.Cells.CreateRange(headerRow, startColumn, 1, columnCount);

            // Apply the bold style to the header row cells of the table
            headerRange.ApplyStyle(boldStyle, styleFlag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
