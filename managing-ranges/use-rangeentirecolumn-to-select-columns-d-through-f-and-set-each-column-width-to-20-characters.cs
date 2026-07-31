// Title: C# Aspose.Cells – Set column width for columns D‑F using Range.EntireColumn
// Description: Creates a workbook, defines the D:F range, accesses its EntireColumn property, sets each column's width to 20 characters, and saves the file as ColumnWidth_D_to_F.xlsx.
// Keywords: Aspose.Cells column width C# | Range.EntireColumn example | set multiple columns width Aspose.Cells | CreateRange D:F | C# Excel column formatting
// Common Searches: Aspose.Cells set width of columns D to F | How to use EntireColumn to change column size in .NET | C# code for setting uniform column width in Excel | Aspose.Cells range column width tutorial | Adjust column width for a range with Aspose.Cells
// Developer Intent: Apply a 20‑character width to columns D, E, and F in a worksheet via the EntireColumn property.
// Use Cases: Standardize column widths in financial statements generated automatically. | Prepare a template workbook where columns D‑F must match a predefined layout. | Export data to Excel with consistent column sizing for better readability.
// AI Prompts: Show a C# snippet that sets the width of columns D through F using Range.EntireColumn in Aspose.Cells. | Explain how to retrieve the current width of a column after modifying it with EntireColumn. | Demonstrate setting different widths for columns D, E, and F in a single Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Creates a workbook, defines the D:F range, accesses its EntireColumn property, sets each column's width to 20 characters, and saves the file as ColumnWidth_D_to_F.xlsx.
public class SetColumnWidthExample
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a range that covers columns D through F
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("D:F");

            // Get the entire columns for the range (columns D, E, F)
            Aspose.Cells.Range entireColumns = range.EntireColumn;

            // Set the width of each column in the range to 20 characters
            entireColumns.ColumnWidth = 20.0;

            // Save the workbook
            string outputPath = "ColumnWidth_D_to_F.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SetColumnWidthExample.Run();
    }
}
