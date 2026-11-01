// Title: Overwrite a cell value in an Excel ListObject table using Cell.PutValue after retrieving it with Cell.GetTable (C# Aspose.Cells)
// AI Prompts: Retrieve the ListObject that contains a specific cell via Cell.GetTable, then assign a new value to the desired data cell using Cell.PutValue and save the workbook. | Change the value of the second row, second column inside a table's DataRange by accessing the ListObject and invoking PutValue on the target cell. | Programmatically update any cell within an Aspose.Cells table: locate the table with GetTable, modify the cell with PutValue, and write the changes to a new file.
// Common Searches: Aspose.Cells C# how to update a value in a ListObject table cell | using Cell.GetTable to modify a specific cell in an Excel table with Aspose | C# replace data in a table row using Aspose.Cells PutValue method | change cell value inside Excel table programmatically with Aspose.Cells .NET | retrieve table from cell A2 and edit its contents Aspose.Cells
// Tags: Cell.GetTable obtain ListObject | Cell.PutValue modify table cell | Aspose.Cells update ListObject data | C# edit Excel table cell | Aspose.Cells overwrite table entry

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a simple workbook with a table if it does not exist, loads the file, uses Cell.GetTable to get the ListObject containing cell A2, updates the value of the second data row and column to "New Value" with Cell.PutValue, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists; create a simple workbook if it doesn't.
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                Worksheet tempWs = tempWb.Worksheets[0];
                tempWs.Cells["A1"].PutValue("Header1");
                tempWs.Cells["B1"].PutValue("Header2");
                tempWs.Cells["A2"].PutValue("Old Value1");
                tempWs.Cells["B2"].PutValue("Old Value2");
                // Create a table covering A1:B2 (hasHeaders = true)
                tempWs.ListObjects.Add(0, 0, 2, 2, true);
                tempWb.Save(inputPath);
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the table that contains cell A2
            ListObject table = sheet.Cells["A2"].GetTable();

            if (table != null)
            {
                // Access a specific cell inside the table (row 2, column 2 of the data range)
                Cell targetCell = table.DataRange[1, 1];
                targetCell.PutValue("New Value");
            }

            // Save the workbook with the updated value
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
