// Title: Insert rows to shift existing data and add a ListObject table at a specific row with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells insert rows before creating a ListObject table | how to shift worksheet rows down when adding a table using Aspose.Cells | Aspose.Cells .NET add ListObject at row index 5 after inserting rows | sample code for inserting rows and creating a table with headers in Aspose.Cells | move existing rows down and place new data table in Excel using Aspose.Cells C#
// Tags: insert rows Aspose.Cells C# | add ListObject table Aspose.Cells | shift worksheet rows before table insertion | populate ListObject with sample data C# | save workbook after row insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, inserts ten rows at zero‑based index 5 to push existing content down, creates a ListObject table spanning columns A‑D with headers, fills the table with sample data, and saves the modified workbook as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define where the new table should start
                int targetFirstRow = 5;          // 0‑based index
                const int rowsToInsert = 10;     // Number of rows for the new table (including header)
                const int totalColumns = 4;      // Number of columns (A‑D)

                // Shift existing rows down to make space for the new table
                sheet.Cells.InsertRows(targetFirstRow, rowsToInsert);

                // Define table dimensions
                int firstColumn = 0; // Column A (0‑based)

                // Add a ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(
                    targetFirstRow,   // FirstRow (0‑based)
                    firstColumn,      // FirstColumn (0‑based)
                    rowsToInsert,     // TotalRows
                    totalColumns,     // TotalColumns
                    true              // HasHeaders
                );
                ListObject table = sheet.ListObjects[tableIndex];
                // Set a display name for the table (Name property may not be available in some versions)
                table.DisplayName = "NewDataTable";

                // Sample data (header + rows)
                string[,] sampleData = new string[rowsToInsert, totalColumns]
                {
                    { "Header1", "Header2", "Header3", "Header4" },
                    { "R1C1",    "R1C2",    "R1C3",    "R1C4" },
                    { "R2C1",    "R2C2",    "R2C3",    "R2C4" },
                    { "R3C1",    "R3C2",    "R3C3",    "R3C4" },
                    { "R4C1",    "R4C2",    "R4C3",    "R4C4" },
                    { "R5C1",    "R5C2",    "R5C3",    "R5C4" },
                    { "R6C1",    "R6C2",    "R6C3",    "R6C4" },
                    { "R7C1",    "R7C2",    "R7C3",    "R7C4" },
                    { "R8C1",    "R8C2",    "R8C3",    "R8C4" },
                    { "R9C1",    "R9C2",    "R9C3",    "R9C4" }
                };

                // Write data into the worksheet cells (the table will reference them automatically)
                for (int i = 0; i < rowsToInsert; i++)
                {
                    for (int j = 0; j < totalColumns; j++)
                    {
                        sheet.Cells[targetFirstRow + i, firstColumn + j].PutValue(sampleData[i, j]);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
