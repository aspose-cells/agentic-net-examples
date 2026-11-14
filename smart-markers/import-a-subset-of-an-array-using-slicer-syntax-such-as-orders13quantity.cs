// Title: How to sum a sliced range of a ListObject column using Aspose.Cells formula syntax in C#
// Common Searches: Aspose.Cells C# formula to sum a specific slice of a table column | Using slicer syntax [1..3] with ListObject in Aspose.Cells | Calculate sum of Orders table rows 2 to 4 Quantity column in C# | Apply range operator 1..3 in Aspose.Cells smart markers example | C# code for summing a subset of a ListObject using Aspose.Cells formulas
// Tags: listobject slicer syntax sum | Aspose.Cells range operator 1..3 | excel formula slice table column c# | smart markers array slicing Aspose.Cells | c# workbook listobject formula example

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example creates a workbook, adds a ListObject named Orders with OrderID and Quantity columns, fills five rows of data, and then uses the slicer syntax (=SUM(Orders[1..3].Quantity)) to sum the Quantity values of rows 2‑4. The result is placed in cell D2 and the workbook is saved as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row for the Orders table
            sheet.Cells["A1"].PutValue("OrderID");
            sheet.Cells["B1"].PutValue("Quantity");

            // Populate sample data (5 rows)
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue(5);
            sheet.Cells["B6"].PutValue(50);

            // Define the range for the ListObject (table) covering A1:B6
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 6;         // rows 1‑6 inclusive
            int totalColumns = 2;      // columns A and B

            // Add the ListObject and retrieve it
            int listObjectIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject ordersTable = sheet.ListObjects[listObjectIndex];
            ordersTable.DisplayName = "Orders";

            // Use slicer syntax to reference a subset of the Quantity column (rows 2‑4, i.e., index 1..3)
            // The formula sums the Quantity values of the selected slice.
            // Place the result in cell D2.
            sheet.Cells["D2"].Formula = "=SUM(Orders[1..3].Quantity)";

            // Save the workbook to a file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
