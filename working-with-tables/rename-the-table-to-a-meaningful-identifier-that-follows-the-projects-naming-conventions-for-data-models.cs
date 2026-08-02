// Title: Rename an Aspose.Cells ListObject (Excel Table) to a Convention‑Compliant Name in C# (.NET)
// Description: Creates a workbook, fills cells A1:C3 with sample order data, adds a ListObject covering the range, sets its DisplayName to a PascalCase identifier (SalesOrdersTable) that follows project naming conventions, and saves the file as RenamedDataModelTable.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells ListObject rename | C# Excel table DisplayName | set ListObject name Aspose.Cells | Excel table naming convention .NET | rename data model table programmatically | Aspose.Cells Table DisplayName C#
// Common Searches: how to rename a ListObject table using Aspose.Cells C# | Aspose.Cells set DisplayName for Excel table | C# change Excel table name with Aspose.Cells | rename data model table in workbook Aspose.Cells | Aspose.Cells ListObject naming conventions
// Developer Intent: Assign a convention‑compliant DisplayName to a ListObject (Excel table) using Aspose.Cells in C#.
// Use Cases: Create a new ListObject and give it a PascalCase name with a "Table" suffix for consistent data‑model references. | Update the DisplayName of an existing table before exporting so downstream processes can locate the correct model. | Standardize table names across multiple worksheets when generating automated reports with Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells that renames an existing ListObject to "CustomerInvoicesTable" following PascalCase rules. | Explain step‑by‑step how to change a ListObject's DisplayName and save the workbook, including proper error handling. | Write a reusable C# method that takes a ListObject and a string, sets the DisplayName, and returns success status.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills cells A1:C3 with sample order data, adds a ListObject covering the range, sets its DisplayName to a PascalCase identifier (SalesOrdersTable) that follows project naming conventions, and saves the file as RenamedDataModelTable.xlsx.
    public class RenameDataModelTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data that will become the table
                worksheet.Cells["A1"].PutValue("OrderID");
                worksheet.Cells["B1"].PutValue("Customer");
                worksheet.Cells["C1"].PutValue("Amount");
                worksheet.Cells["A2"].PutValue(1001);
                worksheet.Cells["B2"].PutValue("John Doe");
                worksheet.Cells["C2"].PutValue(250.75);
                worksheet.Cells["A3"].PutValue(1002);
                worksheet.Cells["B3"].PutValue("Jane Smith");
                worksheet.Cells["C3"].PutValue(180.00);

                // Add a ListObject (Excel table) covering the data range
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 2, true);
                ListObject listObject = worksheet.ListObjects[tableIndex];

                // Assign a meaningful name following the project's naming conventions
                // Example convention: PascalCase with "Table" suffix
                listObject.DisplayName = "SalesOrdersTable";

                // Output the table name to verify
                Console.WriteLine("ListObject (Excel table) DisplayName: " + listObject.DisplayName);

                // Save the workbook to a file
                string outputPath = "RenamedDataModelTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RenameDataModelTableDemo.Run();
        }
    }
}
