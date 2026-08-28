// Title: Create staggered column layout in Excel with Aspose.Cells ImportObjectArray using alternating skip values in C#
// AI Prompts: Generate C# code that uses Aspose.Cells Workbook and Worksheet to import object arrays with a skip of 1 column for the first and third rows and a skip of 0 for the second row, then save the file. | Show how to call Cells.ImportObjectArray with different skip arguments to produce a staggered data pattern across multiple rows. | Provide a step‑by‑step example that creates a workbook, imports three rows with alternating column gaps, and writes the result to an .xlsx file.
// Common Searches: Aspose.Cells C# import object array with column skip parameter | How to create alternating spaced rows in an Excel sheet using Aspose.Cells | ImportObjectArray skip argument for staggered layout in C# | C# Aspose.Cells place data with gaps between columns | Save workbook after importing rows with different skip values Aspose.Cells
// Tags: ImportObjectArray column skip Aspose.Cells | staggered row layout Excel C# | alternating skip parameter worksheet | Aspose.Cells export to XLSX with spaced cells | C# create gap between columns using ImportObjectArray

using System;
using Aspose.Cells;

namespace StaggeredDataLayoutDemo
{
    // Demonstrates using Aspose.Cells Cells.ImportObjectArray in C# to insert rows with and without column gaps, creating a staggered layout and saving the workbook as StaggeredDataLayout.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data to be placed in alternating rows
            // Row 0 will have a skip (blank column) between each entry
            // Row 1 will have no skip (continuous entries)
            object[] rowWithSkip = new object[] { "A1", "B1", "C1", "D1" };
            object[] rowWithoutSkip = new object[] { "A2", "B2", "C2", "D2" };

            // Import first row starting at A1 (row 0, column 0) horizontally with a skip of 1 column
            // This creates a staggered layout: A1, C1, E1, G1 ...
            cells.ImportObjectArray(rowWithSkip, 0, 0, false, 1);

            // Import second row starting at A2 (row 1, column 0) horizontally with no skip (skip = 0)
            // This places data in consecutive cells: A2, B2, C2, D2
            cells.ImportObjectArray(rowWithoutSkip, 1, 0, false, 0);

            // Demonstrate a third row that again uses the skip pattern to continue the alternating effect
            object[] thirdRow = new object[] { "A3", "B3", "C3", "D3" };
            cells.ImportObjectArray(thirdRow, 2, 0, false, 1);

            // Save the workbook to a file
            workbook.Save("StaggeredDataLayout.xlsx");

            Console.WriteLine("Staggered data layout created successfully.");
        }
    }
}
