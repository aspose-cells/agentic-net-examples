// Title: Programmatically reset worksheet print area after inserting rows using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a specified number of rows into a worksheet, compute the new used range, and assign a dynamically built address to PageSetup.PrintArea with Aspose.Cells in C#. | Load an existing Excel file (or create a new workbook), add rows after the current data block, derive the last row and column via MaxDataRow/MaxDataColumn, set the print‑area string, and save the workbook.
// Common Searches: Aspose.Cells C# update print area after adding rows to a worksheet | reset Excel print area programmatically with Aspose.Cells .NET | how to set dynamic print area based on MaxDataRow in Aspose.Cells | C# Aspose.Cells insert rows and adjust page setup print area | change print area range after expanding data in an Aspose.Cells workbook
// Tags: dynamic print area Aspose.Cells | worksheet row insertion C# | reset print area after data expansion .NET | PageSetup.PrintArea assignment Aspose.Cells | save workbook after print area update

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads (or creates) an Excel workbook, accesses the first worksheet, inserts five rows after the existing data, recalculates the last used row and column, builds a new A1:... address, assigns it to worksheet.PageSetup.PrintArea, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; if not, create an empty workbook.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Access the first worksheet (modify as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Example modification: insert 5 new rows after the current data range
                int rowsToAdd = 5;
                int insertPosition = worksheet.Cells.MaxDataRow + 1; // position after last used row
                worksheet.Cells.InsertRows(insertPosition, rowsToAdd);

                // Re‑calculate the new data bounds
                int lastRow = worksheet.Cells.MaxDataRow;               // zero‑based index
                int lastColumn = worksheet.Cells.MaxDataColumn;         // zero‑based index

                // Build the address of the new print area (e.g., A1:D20)
                string startCell = "A1";
                string endCell = CellsHelper.ColumnIndexToName(lastColumn) + (lastRow + 1).ToString(); // rows are 1‑based in addresses
                worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
