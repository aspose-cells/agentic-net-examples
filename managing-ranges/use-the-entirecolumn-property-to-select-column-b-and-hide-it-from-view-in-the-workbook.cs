// Title: C# – Hide Column B in an Aspose.Cells Workbook via Range.EntireColumn
// Description: Demonstrates how to create a workbook, insert data into column B, obtain the full column using the Range.EntireColumn accessor, and conceal it with Cells.HideColumn before saving the file.
// Keywords: Aspose.Cells hide column C# | Range.EntireColumn example | programmatically hide Excel column | Aspose.Cells column visibility | C# Excel column conceal
// Common Searches: Aspose.Cells hide specific column using EntireColumn | C# code to conceal column B in Excel workbook | How to use Range.EntireColumn to hide columns in Aspose.Cells
// Developer Intent: Conceal column B in a generated workbook by selecting it through the EntireColumn accessor.
// Use Cases: Prevent sensitive data in a column from being displayed in shared reports. | Dynamically hide columns based on user preferences in a dashboard application. | Batch‑process a list of columns, using Range.EntireColumn to toggle their visibility before export.
// AI Prompts: Write C# that hides columns C through E in an Aspose.Cells workbook using Range.EntireColumn. | Show how to toggle a column's hidden state at runtime with a boolean flag in Aspose.Cells for .NET. | Explain the steps to iterate over multiple column letters and hide each using the EntireColumn property.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, insert data into column B, obtain the full column using the Range.EntireColumn accessor, and conceal it with Cells.HideColumn before saving the file.
class HideColumnUsingEntireColumn
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Add sample data to column B
            cells["B1"].PutValue("This column will be hidden");
            cells["B2"].PutValue(123.45);

            // Select column B using the EntireColumn property of a range that starts at B1
            AsposeRange range = cells.CreateRange("B1");          // Create a range from cell B1
            AsposeRange entireColumn = range.EntireColumn;        // Represents the whole column B

            // Hide column B (zero‑based index 1)
            cells.HideColumn(1);

            // Determine output file path and ensure directory exists
            string outputFile = "HideColumnUsingEntireColumn.xlsx";
            string fullPath = Path.GetFullPath(outputFile);
            string dir = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // Save the workbook
            workbook.Save(fullPath);
            Console.WriteLine($"Workbook saved to: {fullPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
