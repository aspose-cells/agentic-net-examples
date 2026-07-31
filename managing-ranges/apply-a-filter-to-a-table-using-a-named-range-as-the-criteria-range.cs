// Title: Apply an In‑Place Advanced Filter Using a Named Criteria Range in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates a product table, defines a named range (MyCriteria) with a header and value, and runs Worksheet.AdvancedFilter to filter rows where Category = "Fruit" directly on the source range A1:C5, then saves the result.
// Keywords: Aspose.Cells AdvancedFilter C# | named range criteria Aspose.Cells | filter table by category .NET | in‑place worksheet filter | Aspose.Cells example filtering
// Common Searches: Aspose.Cells use named range for AdvancedFilter | C# filter Excel table with criteria range | how to apply AdvancedFilter without copying data | filter rows by category using Aspose.Cells
// Developer Intent: Use a predefined named range as the criteria source to filter a worksheet table with Aspose.Cells' AdvancedFilter method.
// Use Cases: Display only fruit items by setting MyCriteria to Category = "Fruit". | Reuse the same named criteria to filter several tables in one workbook. | Apply the filter in place to keep the original layout while hiding non‑matching rows.
// AI Prompts: Write C# code that creates a named range and applies Worksheet.AdvancedFilter with it in Aspose.Cells. | Explain step‑by‑step how to filter an Excel table in place using a named criteria range in Aspose.Cells for .NET. | Show how to change the value in a named criteria range at runtime to filter different categories.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates a product table, defines a named range (MyCriteria) with a header and value, and runs Worksheet.AdvancedFilter to filter rows where Category = "Fruit" directly on the source range A1:C5, then saves the result.
    public class FilterTableWithNamedCriteriaDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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

            // ---------- Populate sample data for the table ----------
            // Header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Category");
            worksheet.Cells["C1"].PutValue("Price");

            // Data rows
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue("Fruit");
            worksheet.Cells["C2"].PutValue(1.2);

            worksheet.Cells["A3"].PutValue("Carrot");
            worksheet.Cells["B3"].PutValue("Vegetable");
            worksheet.Cells["C3"].PutValue(0.8);

            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B4"].PutValue("Fruit");
            worksheet.Cells["C4"].PutValue(1.0);

            worksheet.Cells["A5"].PutValue("Broccoli");
            worksheet.Cells["B5"].PutValue("Vegetable");
            worksheet.Cells["C5"].PutValue(1.5);

            // ---------- Create a named range that will serve as the criteria ----------
            // Criteria header (must match the column header in the list range)
            worksheet.Cells["E1"].PutValue("Category");
            // Criteria value (e.g., filter to show only rows where Category = "Fruit")
            worksheet.Cells["E2"].PutValue("Fruit");

            // Define the range E1:E2 and assign a name
            AsposeRange criteriaRange = worksheet.Cells.CreateRange("E1:E2");
            criteriaRange.Name = "MyCriteria";

            // ---------- Apply an advanced filter using the named criteria range ----------
            // List range address (including header row)
            string listRange = "A1:C5";

            // The criteria range can be referenced by its name
            string criteriaRangeName = "MyCriteria";

            // Apply the filter in place (isFilter = true), no copy destination, unique records not required
            worksheet.AdvancedFilter(true, listRange, criteriaRangeName, null, false);

            // ---------- Save the workbook ----------
            string outputPath = "FilterTableWithNamedCriteriaDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}
