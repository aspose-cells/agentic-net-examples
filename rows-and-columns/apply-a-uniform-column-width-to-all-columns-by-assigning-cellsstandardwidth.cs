// Title: Apply Uniform Column Width to All Columns with Aspose.Cells Cells.StandardWidth (C#)
// Description: Demonstrates how to set a single column width for every column in an Excel worksheet using the Aspose.Cells Cells.StandardWidth property in C#. The example creates a workbook, assigns a uniform width (character units), verifies the setting, and saves the file as UniformColumnWidth.xlsx.
// Keywords: Aspose.Cells set column width C# | Cells.StandardWidth property | uniform column width Aspose.Cells | apply same width to all columns | Excel column width character units | Aspose.Cells column formatting
// Common Searches: how to set same column width for all columns Aspose.Cells | Aspose.Cells Cells.StandardWidth example | C# set uniform column width Excel file | retrieve column width after setting standard width Aspose | Aspose.Cells column width tutorial
// Developer Intent: Set an identical width for every column in a worksheet and save the workbook.
// Use Cases: Create a fresh workbook and enforce a 20‑character column width before inserting data. | Open an existing spreadsheet, standardize column widths across the sheet with Cells.StandardWidth, then save the changes. | After applying Cells.StandardWidth, read the actual width of a specific column to confirm the adjustment and log the result.
// AI Prompts: Generate C# code that uses Aspose.Cells to set Cells.StandardWidth to 25, adds a header row, populates sample data, and saves the workbook as 'Report.xlsx'. | Explain the effect of the Cells.StandardWidth property on column sizing and show how to retrieve the actual width of column 0 after setting it.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to set a single column width for every column in an Excel worksheet using the Aspose.Cells Cells.StandardWidth property in C#. The example creates a workbook, assigns a uniform width (character units), verifies the setting, and saves the file as UniformColumnWidth.xlsx.
    public class ApplyUniformColumnWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default worksheet is added)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the Cells collection of the worksheet
                Cells cells = worksheet.Cells;

                // Set a uniform column width for all columns in the worksheet.
                // This value is in character units (the width of the digit '0' in the default font).
                cells.StandardWidth = 20.0; // Adjust the value as needed

                // Optional: Verify that the standard width has been applied to a specific column
                Console.WriteLine("Standard Width set to: " + cells.StandardWidth);
                Console.WriteLine("Actual width of column 0: " + cells.GetColumnWidth(0));

                // Define output file path
                string outputPath = "UniformColumnWidth.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
