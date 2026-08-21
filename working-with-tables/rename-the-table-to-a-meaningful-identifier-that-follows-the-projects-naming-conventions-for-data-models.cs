// Title: Rename an Aspose.Cells ListObject (Excel Table) Using the DisplayName Property – C# Example
// Description: Demonstrates how to create a workbook with Aspose.Cells for .NET, add a ListObject covering A1:B3, assign a meaningful DisplayName such as "SalesData", and save the file. Shows the programmatic way to rename Excel tables to match project naming conventions.
// Keywords: Aspose.Cells rename ListObject | C# Excel table DisplayName | change Excel table name programmatically | Aspose.Cells set table identifier | rename Aspose.Cells table .NET
// Common Searches: how to rename a ListObject in Aspose.Cells C# | set DisplayName for Excel table using Aspose.Cells | Aspose.Cells change table name .NET | rename Excel table programmatically Aspose
// Developer Intent: Rename a ListObject to a clear, convention‑compliant identifier within an Aspose.Cells workbook.
// Use Cases: Align generated table names with data‑model entities (e.g., SalesData) before distribution. | Enforce a consistent naming scheme across multiple worksheets for downstream analytics. | Dynamically adjust table identifiers based on user input, configuration files, or localization.
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, adds a ListObject, and renames the table to "CustomerOrders" using the DisplayName property. | Provide a C# snippet that iterates over all ListObjects in a workbook and renames each according to a pattern like "Tbl_{SheetName}_{Index}" with Aspose.Cells. | Explain how to validate a proposed table name against a naming convention (e.g., PascalCase, max length) before assigning it to ListObject.DisplayName in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with Aspose.Cells for .NET, add a ListObject covering A1:B3, assign a meaningful DisplayName such as "SalesData", and save the file. Shows the programmatic way to rename Excel tables to match project naming conventions.
    public class RenameTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Quantity");
                worksheet.Cells["A2"].PutValue("Widget");
                worksheet.Cells["B2"].PutValue(150);
                worksheet.Cells["A3"].PutValue("Gadget");
                worksheet.Cells["B3"].PutValue(85);

                // Add a ListObject (table) covering the data range A1:B3
                int tableIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Rename the table
                table.DisplayName = "SalesData";

                // Save the workbook
                string outputPath = "RenamedTableDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            RenameTableDemo.Run();
        }
    }
}
