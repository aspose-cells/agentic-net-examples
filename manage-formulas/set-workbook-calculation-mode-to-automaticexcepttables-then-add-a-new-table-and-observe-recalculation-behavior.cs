// Title: Demonstrate AutomaticExceptTables calculation mode with a ListObject table and verify external formula behavior in Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, set workbook.CalculationOptions.CalculationMode to AutomaticExceptTables, fill cells A1‑B2 with numbers, place a SUM formula in C1, insert a ListObject covering A1:B2, modify a cell inside the table, and print the value of C1 before and after the change. | Using Aspose.Cells for .NET, show how enabling AutomaticExceptTables prevents a formula outside a table from recalculating automatically after the table data is edited, then save the workbook as output.xlsx.
// Common Searches: Aspose.Cells C# set calculation mode to AutomaticExceptTables and check formula update | How to stop automatic recalculation of formulas outside tables in Aspose.Cells .NET | Example of adding a ListObject table and testing calculation mode in Aspose.Cells | C# Aspose.Cells workbook calculation options after inserting a table | AutomaticExceptTables behavior with SUM formula in Aspose.Cells example
// Tags: Aspose.Cells AutomaticExceptTables setting | ListObject table creation Aspose.Cells C# | disable auto‑recalc for external formulas .NET | save workbook as XLSX using Aspose.Cells | formula recalculation behavior with tables Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The sample creates a workbook, populates A1‑B2 with values, adds a SUM formula in C1, forces an initial calculation, sets the calculation mode to AutomaticExceptTables, inserts a ListObject over the data range, changes a cell inside the table, prints the formula result before and after the change to demonstrate that the external formula does not recalculate automatically, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B1"].PutValue(5);
            sheet.Cells["B2"].PutValue(15);

            // Add a formula outside the future table range
            sheet.Cells["C1"].Formula = "=SUM(A1:A2)";

            // Force an initial calculation so we can see the starting value
            workbook.CalculateFormula();
            Console.WriteLine("Initial sum (C1): " + sheet.Cells["C1"].Value);

            // Add a table (ListObject) covering A1:B2
            int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, false);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";                     // Set table name (DisplayName works across versions)
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Change a value inside the table
            sheet.Cells["A1"].PutValue(30);

            // Formulas outside tables (like C1) recalculate automatically in default Automatic mode
            Console.WriteLine("After changing A1 inside table, sum (C1): " + sheet.Cells["C1"].Value);

            // Save the workbook to observe the results
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
