// Title: Aspose.Cells .NET: Set width of columns D‑F to 20 characters using Range.EntireColumn
// Description: Creates a workbook, defines the range D:F, accesses its EntireColumn property, sets each column's width to 20 characters, and saves the file as Columns_D_to_F_Width_20.xlsx.
// Keywords: Aspose.Cells | .NET | C# | Range EntireColumn | set column width | columns D-F | Excel column sizing | CreateRange D:F
// Common Searches: Aspose.Cells set column width D to F | Range.EntireColumn column width C# example | How to change multiple column widths with Aspose.Cells | C# Excel column width 20 characters Aspose
// Developer Intent: Apply a uniform width of 20 characters to columns D, E, and F in an Excel workbook via Aspose.Cells.
// Use Cases: Formatting report columns for consistent appearance. | Preparing header rows before data insertion to ensure alignment. | Reusing column‑width logic across several worksheets in automated Excel generation.
// AI Prompts: Write C# code that sets columns G‑I to a width of 25 characters using Aspose.Cells Range.EntireColumn. | Compare Range.EntireColumn.ColumnWidth with setting Cells.ColumnWidth individually for multiple columns in Aspose.Cells. | Provide a step‑by‑step tutorial to create a range for columns A‑C and set their width to 15 characters with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, defines the range D:F, accesses its EntireColumn property, sets each column's width to 20 characters, and saves the file as Columns_D_to_F_Width_20.xlsx.
    public class SetColumnWidthUsingEntireColumn
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Create a range that covers columns D through F (columns 3 to 5, zero‑based)
                Aspose.Cells.Range range = worksheet.Cells.CreateRange("D:F");

                // Get the entire columns for the range
                Aspose.Cells.Range entireColumns = range.EntireColumn;

                // Set the width of each column in the range to 20 characters
                entireColumns.ColumnWidth = 20.0;

                // Save the workbook
                string outputPath = "Columns_D_to_F_Width_20.xlsx";
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
            SetColumnWidthUsingEntireColumn.Run();
        }
    }
}
