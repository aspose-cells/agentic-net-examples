// Title: Detect Null or Missing Cells in an Aspose.Cells Worksheet (.NET)
// Description: Creates a workbook, populates sample data, determines the used range with Cells.MaxDataRow/MaxDataColumn, iterates each cell using Cells.CheckCell, logs addresses of cells that are absent or have a null value, and saves the file. Ideal for data‑quality checks in C# projects.
// Keywords: Aspose.Cells null cell detection | C# enumerate worksheet cells | CheckCell missing cell Aspose | log empty cells Aspose.Cells | CellsHelper.CellIndexToName example | used range iteration Aspose.Cells
// Common Searches: how to find empty cells with Aspose.Cells .NET | Aspose.Cells iterate used range and detect null values | log addresses of missing cells Aspose.Cells | CheckCell vs GetCell null handling in Aspose | C# Aspose.Cells data completeness audit
// Developer Intent: Identify and record every blank or non‑existent cell while looping through a worksheet’s used area.
// Use Cases: Perform a data‑quality audit that lists all blank cells before exporting a spreadsheet. | Skip null cells during bulk calculations to reduce processing time. | Generate a report of missing values for downstream validation or ETL pipelines.
// AI Prompts: Write C# code using Aspose.Cells that collects null cell addresses into a List<string> instead of printing them. | Show how to modify the loop to stop at the first null cell and throw a custom DataIntegrityException. | Explain how Cells.MaxDataRow and Cells.MaxDataColumn define safe iteration bounds for large worksheets when detecting empty cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates sample data, determines the used range with Cells.MaxDataRow/MaxDataColumn, iterates each cell using Cells.CheckCell, logs addresses of cells that are absent or have a null value, and saves the file. Ideal for data‑quality checks in C# projects.
    public class DetectNullCellsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data with some empty (null) cells
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["A2"].PutValue(100);
                // B2 is left empty intentionally (null)

                // Determine the used range boundaries
                int maxRow = cells.MaxDataRow;       // zero‑based index of the last row with data
                int maxColumn = cells.MaxDataColumn; // zero‑based index of the last column with data

                // Iterate through the used range and log cells that are null or missing
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // CheckCell returns null if the cell does not exist
                        Cell cell = cells.CheckCell(row, col);

                        // If the cell is missing or its value is null, log its address
                        if (cell == null || cell.Value == null)
                        {
                            string address = CellsHelper.CellIndexToName(row, col);
                            Console.WriteLine($"Null or missing cell detected at {address}");
                        }
                    }
                }

                // Save the workbook (the file will contain the sample data)
                string outputPath = "DetectNullCellsDemo.xlsx";
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
            DetectNullCellsDemo.Run();
        }
    }
}
