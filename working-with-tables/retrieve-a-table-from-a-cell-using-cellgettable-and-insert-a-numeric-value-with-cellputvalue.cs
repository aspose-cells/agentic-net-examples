// Title: Retrieve a ListObject from a cell with Cell.GetTable and write a numeric value using Cell.PutValue in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get ListObject from a specific cell | C# example using Cell.GetTable to access a worksheet table | How to insert a number into a cell after retrieving a table with Aspose.Cells | Aspose.Cells put numeric value into cell C2 after GetTable | Retrieve table from cell A2 and write value 12345 using Aspose.Cells .NET
// Tags: Cell.GetTable API usage | Cell.PutValue numeric insertion | Aspose.Cells ListObject creation range A1:B3 | Aspose.Cells workbook save to XLSX | C# retrieve table from cell

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, fills cells A1:B3 with headers and numbers, adds a ListObject covering that range, retrieves the table via Cell.GetTable from cell A2, and if the table is found writes the value 12345 to cell C2 using Cell.PutValue before saving the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data to form a table
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B3"].PutValue(40);

            // Add a table (list object) covering the range A1:B3
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Set a display name for the table (Name property may not be available in some versions)
            table.DisplayName = "MyTable";

            // Retrieve the table from a cell that belongs to the table
            Cell cellInTable = sheet.Cells["A2"];
            ListObject retrievedTable = cellInTable.GetTable();

            // If the table is successfully retrieved, insert a numeric value into another cell
            if (retrievedTable != null)
            {
                // Insert numeric value using PutValue (insert rule)
                sheet.Cells["C2"].PutValue(12345);
            }

            // Determine output path and ensure its directory exists
            string outputPath = "Result.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
