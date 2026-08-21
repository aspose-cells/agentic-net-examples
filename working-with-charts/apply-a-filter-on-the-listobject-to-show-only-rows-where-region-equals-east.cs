// Title: C# – Filter a ListObject (Table) by Region = "East" with Aspose.Cells AutoFilter
// Description: Demonstrates how to create a workbook, add a ListObject, enable AutoFilter, and programmatically show only rows where the "Region" column equals "East" using Aspose.Cells for .NET.
// Keywords: Aspose.Cells ListObject filter C# | AutoFilter table Aspose.Cells | filter rows by column value .NET | Region column filter Aspose.Cells | C# Aspose.Cells table example | ListObject AutoFilter Refresh
// Common Searches: Aspose.Cells filter ListObject by column value | C# AutoFilter on Aspose.Cells table | How to show only rows where Region = East in Aspose.Cells | ListObject.AutoFilter.Filter example C# | Apply column filter to Aspose.Cells worksheet
// Developer Intent: Programmatically display only the rows in a ListObject where the Region column equals "East".
// Use Cases: Generate regional sales reports that automatically hide non‑East data before export. | Build an interactive dashboard where selecting a region applies an AutoFilter to the underlying table. | Pre‑process large datasets by isolating specific geographic subsets using ListObject filters.
// AI Prompts: Write C# code with Aspose.Cells to add a ListObject, enable AutoFilter, and filter rows where the "Region" column equals "East". | Explain the purpose of ListObject.AutoFilter.Filter and Refresh methods in Aspose.Cells and how to use them for different columns. | Provide a step‑by‑step tutorial for creating a table, turning on AutoFilter, applying a value filter, and saving the workbook in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a ListObject, enable AutoFilter, and programmatically show only rows where the "Region" column equals "East" using Aspose.Cells for .NET.
    public class ListObjectRegionFilterDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a "Region" column
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Region");
            worksheet.Cells["A2"].PutValue("Laptop");
            worksheet.Cells["B2"].PutValue("East");
            worksheet.Cells["A3"].PutValue("Smartphone");
            worksheet.Cells["B3"].PutValue("West");
            worksheet.Cells["A4"].PutValue("Monitor");
            worksheet.Cells["B4"].PutValue("East");
            worksheet.Cells["A5"].PutValue("Tablet");
            worksheet.Cells["B5"].PutValue("North");

            // Create a ListObject (table) that covers the data range (including header)
            // Parameters: startRow, startColumn, endRow, endColumn, hasHeaders
            int listObjectIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject listObject = worksheet.ListObjects[listObjectIndex];

            // Ensure the table has an AutoFilter enabled
            listObject.HasAutoFilter = true;

            // Apply a filter on the "Region" column (index 1) to show only rows where Region = "East"
            listObject.AutoFilter.Filter(1, "East");
            listObject.AutoFilter.Refresh();

            // Save the workbook
            string outputPath = "ListObjectRegionFilterDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
