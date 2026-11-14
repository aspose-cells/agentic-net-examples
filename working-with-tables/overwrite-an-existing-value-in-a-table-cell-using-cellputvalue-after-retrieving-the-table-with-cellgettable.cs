// Title: Update a specific Excel table cell using Cell.PutValue after locating its ListObject with Cell.GetTable in Aspose.Cells for .NET
// Common Searches: how to change a cell value inside an Excel table using Aspose.Cells C# | using Cell.GetTable to find the ListObject and then update the cell with PutValue | replace existing data in a ListObject row cell programmatically with Aspose.Cells | Aspose.Cells C# overwrite value of a table cell after checking if it belongs to a table | C# sample to modify a table cell only when the cell is part of a ListObject
// Tags: Cell.GetTable retrieve ListObject | Cell.PutValue update table cell | Aspose.Cells modify Excel table cell | C# overwrite ListObject cell value | check cell belongs to table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an existing workbook, accesses cell B2, uses Cell.GetTable to obtain the containing ListObject, overwrites the cell's value with Cell.PutValue when the cell is part of a table, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Choose a cell that may be part of a table (e.g., B2)
            Cell cell = worksheet.Cells["B2"];

            // Retrieve the table (ListObject) that contains this cell
            ListObject table = cell.GetTable();

            // If the cell belongs to a table, overwrite its value
            if (table != null)
            {
                cell.PutValue("Updated Value");
            }
            else
            {
                Console.WriteLine("Cell B2 is not part of any table.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
