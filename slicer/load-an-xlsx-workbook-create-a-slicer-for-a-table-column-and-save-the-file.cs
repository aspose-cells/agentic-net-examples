// Title: C# example: load an XLSX workbook, add a slicer to a table column, and save the file with Aspose.Cells
// Common Searches: how to programmatically add a slicer to an Excel table using Aspose.Cells in C# | Aspose.Cells C# code to create a slicer for the first column of a ListObject | sample C# program that loads a workbook, inserts a slicer at a specific cell, and saves it | using Aspose.Cells to insert a slicer into an existing XLSX file while preserving formatting | C# tutorial for adding an Excel slicer with Aspose.Cells and exporting as XLSX
// Tags: Aspose.Cells add slicer to ListObject C# | C# create Excel slicer at cell position | Aspose.Cells load workbook insert slicer | save workbook with slicer Aspose.Cells Xlsx | Excel table slicer generation using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

// This C# example demonstrates how to load an existing XLSX workbook (or create one with sample data), ensure a ListObject table covers the data, add a slicer for the table's first column at cell E2, and save the modified workbook as a new XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output_with_slicer.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data
            if (File.Exists(inputPath))
            {
                Console.WriteLine($"Loading workbook from \"{inputPath}\"...");
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook with sample data.");
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];

                // Populate sample data for the table
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(20);
                ws.Cells["A4"].PutValue("A");
                ws.Cells["B4"].PutValue(30);
                ws.Cells["A5"].PutValue("C");
                ws.Cells["B5"].PutValue(40);
            }

            // Work with the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is a table (ListObject) covering the data range.
            // If a table already exists, reuse it; otherwise create a new one.
            ListObject table;
            if (worksheet.ListObjects.Count > 0)
            {
                table = worksheet.ListObjects[0];
                Console.WriteLine($"Using existing table \"{table.DisplayName}\".");
            }
            else
            {
                // Determine the used range to define the table boundaries
                int firstRow = worksheet.Cells.MinRow;
                int firstColumn = worksheet.Cells.MinColumn;
                int lastRow = worksheet.Cells.MaxRow;
                int lastColumn = worksheet.Cells.MaxColumn;

                // Add a table that includes the header row
                int tableIndex = worksheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
                table = worksheet.ListObjects[tableIndex];
                Console.WriteLine($"Created new table \"{table.DisplayName}\" covering range {firstRow},{firstColumn}:{lastRow},{lastColumn}.");
            }

            // Choose the column for which the slicer will be created (first column of the table)
            ListColumn slicerColumn = table.ListColumns[0];

            // Define the position where the slicer will be placed (upper-left cell)
            string slicerPosition = "E2";

            // Add the slicer to the worksheet
            int slicerIndex = worksheet.Slicers.Add(table, slicerColumn, slicerPosition);
            Slicer slicer = worksheet.Slicers[slicerIndex];

            Console.WriteLine($"Slicer added. Table: \"{table.DisplayName}\", Column: \"{slicerColumn.Name}\", Position: {slicerPosition}");

            // Optional: customize slicer style (e.g., set a built‑in style)
            // slicer.Style = SlicerStyleType.SlicerStyleLight1;

            // Save the workbook with the slicer
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
        }
    }
}
