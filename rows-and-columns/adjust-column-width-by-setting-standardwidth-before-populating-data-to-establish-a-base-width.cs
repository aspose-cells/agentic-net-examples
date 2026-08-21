// Title: C# Example: Set Global Column Width with Cells.StandardWidth in Aspose.Cells before Adding Data
// Description: Demonstrates how to define a uniform default column width using the Cells.StandardWidth property, populate sample cells, retrieve the actual width of a column, and save the workbook as AdjustedStandardWidth.xlsx.
// Keywords: Aspose.Cells C# column width | Cells.StandardWidth .NET | set default column width Aspose | global column width before data | Aspose.Cells sample code | adjust column width programmatically | retrieve column width Aspose.Cells
// Common Searches: how to set default column width in Aspose.Cells C# | Cells.StandardWidth example | Aspose.Cells set column width for all columns | get actual column width after setting StandardWidth | C# Aspose.Cells column width before inserting data
// Developer Intent: Define a base column width for the entire worksheet prior to inserting any cell values.
// Use Cases: Create a template where every column starts with a predefined width, ensuring consistent layout across new worksheets. | Compare the globally set StandardWidth with the measured width of a specific column after data entry to validate sizing. | Apply a uniform column width before populating large datasets, then fine‑tune individual columns as needed.
// AI Prompts: Write C# code that sets Cells.StandardWidth to 25 characters, adds a header row, and autosizes selected columns using Aspose.Cells. | Explain how Cells.StandardWidth interacts with GetColumnWidth and column‑specific width overrides in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for setting a default column width, inserting sample data, and verifying column sizes with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to define a uniform default column width using the Cells.StandardWidth property, populate sample cells, retrieve the actual width of a column, and save the workbook as AdjustedStandardWidth.xlsx.
    public class AdjustColumnWidthUsingStandardWidth
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Set the default column width (in character units) before adding any data
            // This establishes a base width for all columns
            cells.StandardWidth = 20.0;

            // Populate sample data to demonstrate the effect of the standard width
            cells["A1"].PutValue("Short");
            cells["B1"].PutValue("This is a longer piece of text that may need more width");
            cells["C1"].PutValue("Medium length");

            // Output the standard width and the actual width of the first column
            Console.WriteLine("Standard Width set to: " + cells.StandardWidth);
            Console.WriteLine("Column A actual width: " + cells.GetColumnWidth(0));

            // Save the workbook to a file
            workbook.Save("AdjustedStandardWidth.xlsx");
        }
    }
}
