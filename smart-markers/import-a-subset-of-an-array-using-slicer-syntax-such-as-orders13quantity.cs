// Title: Use slicer syntax to import a range of rows from a ListObject column in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, defines a ListObject named Orders, and sets the formula =Orders[1..3].Quantity in cell D2 to extract the first three quantities. | Create a reusable method that accepts start and end indices, builds a slicer expression for a ListObject column, writes the result to a target range on any worksheet, and saves the workbook. | Show how to pull data from one ListObject into another worksheet by applying a dynamic slicer range, then demonstrate saving the updated workbook.
// Common Searches: asp.net aspose.cells how to use slicer syntax to get rows 2 to 5 from a table column | c# retrieve specific rows from an Excel table using Aspose.Cells | aspose.cells example extracting subset of a ListObject column in .NET | dynamic range operator for ListObject data with Aspose.Cells C# | import selected rows of a table column into another sheet using Aspose.Cells
// Tags: Aspose.Cells ListObject range extraction | C# import subset of table column | Excel slicer syntax .NET | dynamic range operator Aspose.Cells | extract table rows using Aspose.Cells API

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The program creates a workbook, adds a ListObject named Orders with sample data, applies the slicer formula =Orders[1..3].Quantity to pull the Quantity values of rows 1‑3 into cell D2, and saves the file as OrdersSubset.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the Orders table
            // Header row
            sheet.Cells["A1"].PutValue("OrderID");
            sheet.Cells["B1"].PutValue("Quantity");

            // Data rows
            sheet.Cells["A2"].PutValue(101);
            sheet.Cells["B2"].PutValue(5);
            sheet.Cells["A3"].PutValue(102);
            sheet.Cells["B3"].PutValue(8);
            sheet.Cells["A4"].PutValue(103);
            sheet.Cells["B4"].PutValue(12);
            sheet.Cells["A5"].PutValue(104);
            sheet.Cells["B5"].PutValue(7);

            // Define a ListObject (table) named "Orders" covering the data range
            int firstRow = 0;      // zero‑based index (A1)
            int firstColumn = 0;
            int totalRows = 5;     // includes header
            int totalColumns = 2;

            // Add the table and retrieve the ListObject instance
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);
            ListObject ordersTable = sheet.ListObjects[tableIndex];
            ordersTable.DisplayName = "Orders";

            // Use slicer syntax to import a subset of the Quantity column (rows 1 to 3)
            // Place the result starting at cell D2
            // The formula uses the slicer syntax: =Orders[1..3].Quantity
            sheet.Cells["D2"].Formula = "=Orders[1..3].Quantity";

            // Save the workbook to a file
            string outputPath = "OrdersSubset.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
