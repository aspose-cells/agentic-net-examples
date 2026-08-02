// Title: C# – Set Worksheet Default Column Width to 15 Characters with Aspose.Cells
// Description: This C# snippet creates a new Workbook, selects the first Worksheet, assigns a 15‑character width to all columns via the Cells.StandardWidth property, prints the applied value, and saves the file as DefaultColumnWidth.xlsx.
// Keywords: Aspose.Cells | C# | StandardWidth | worksheet column width | set column width 15 | Excel export | save workbook | column width property
// Common Searches: Aspose.Cells set default column width | How to change column width for all columns in .NET | StandardWidth property example C# | Set column width to 15 characters Aspose.Cells | C# worksheet column width Aspose.Cells
// Developer Intent: Apply a 15‑character width to every column in a worksheet and persist the workbook.
// Use Cases: Initialize a fresh workbook and enforce a uniform column width before populating data. | Adjust the column width of an existing sheet to maintain consistent layout when generating Excel reports. | Validate the applied width by reading Cells.StandardWidth and the actual width of a specific column.
// AI Prompts: Write C# code using Aspose.Cells that sets Cells.StandardWidth to 15 for a worksheet and saves the workbook. | Explain the impact of the StandardWidth property on column sizing and how to retrieve the actual width of a column after setting it. | Provide a step‑by‑step guide to change the default column width for all worksheets in a workbook with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# snippet creates a new Workbook, selects the first Worksheet, assigns a 15‑character width to all columns via the Cells.StandardWidth property, prints the applied value, and saves the file as DefaultColumnWidth.xlsx.
    public class SetDefaultColumnWidth
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the default column width for the entire worksheet to 15 characters
                worksheet.Cells.StandardWidth = 15.0;

                // Optional: verify the setting
                Console.WriteLine("Standard Width set to: " + worksheet.Cells.StandardWidth);
                Console.WriteLine("Column 0 actual width: " + worksheet.Cells.GetColumnWidth(0));

                // Determine output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "DefaultColumnWidth.xlsx");

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine("Workbook saved to: " + outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDefaultColumnWidth.Run();
        }
    }
}
