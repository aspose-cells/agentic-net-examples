// Title: Configure Aspose.Cells smart markers in C# to automatically ignore rows with empty values
// AI Prompts: Write C# code that defines a smart marker range, binds a list containing empty strings, processes it with WorkbookDesigner, and then removes completely blank rows using DeleteBlankOptions. | Demonstrate how to set EmptyStringAsBlank and UpdateReference in DeleteBlankRows to clean up gaps after smart marker processing. | Modify the example so that rows with any null or empty property are not inserted at all, eliminating the need for post‑processing deletion.
// Common Searches: Aspose.Cells C# smart markers ignore rows with empty fields | How to delete blank rows after WorkbookDesigner.Process in .NET | Treat empty string cells as blanks when using DeleteBlankRows Aspose.Cells | Skip null or empty values in smart marker detail list C# example | Configure smart marker options to prevent blank rows in generated Excel file
// Tags: smart markers delete blank rows Aspose.Cells | WorkbookDesigner EmptyStringAsBlank option | C# ignore empty values in smart marker list | Aspose.Cells DeleteBlankRows configuration | smart marker range processing with gaps

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

namespace SmartMarkerIgnoreEmptyRowsDemo
{
    // Sample data class for the detail list
    // The sample creates a workbook, sets up smart markers for a product list, supplies a list that includes empty strings, processes the markers with WorkbookDesigner, and then uses DeleteBlankOptions (EmptyStringAsBlank and UpdateReference) to delete rows that become completely blank before saving the file.
    public class Product
    {
        public string? Name { get; set; }
        public string? Price { get; set; }   // Nullable to allow empty values without warnings
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (template)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // 2. Set up smart markers for a detail list (range will be processed)
                // Header row
                sheet.Cells["A1"].PutValue("Product Name");
                sheet.Cells["B1"].PutValue("Price");

                // Detail row with smart markers
                sheet.Cells["A2"].PutValue("&=Products.Name");
                sheet.Cells["B2"].PutValue("&=Products.Price");

                // Define the range that contains the smart markers and give it the required name
                AsposeRange smartRange = sheet.Cells.CreateRange("A2:B2");
                smartRange.Name = "_CellsSmartMarkers";

                // 3. Prepare a list that contains gaps (empty strings)
                var products = new List<Product>
                {
                    new Product { Name = "Laptop", Price = "1200" },
                    new Product { Name = "", Price = "800" },          // Empty Name – should be ignored
                    new Product { Name = "Tablet", Price = "" },      // Empty Price – should be ignored
                    new Product { Name = "Smartphone", Price = "600" }
                };

                // 4. Configure the WorkbookDesigner
                var designer = new WorkbookDesigner
                {
                    Workbook = workbook
                    // LineByLine is obsolete; default behavior works with range smart markers
                };

                // 5. Set the data source for the smart markers
                designer.SetDataSource("Products", products);

                // 6. Process the smart markers – this will populate rows for each item in the list
                designer.Process();

                // 7. After processing, delete rows that are completely blank (empty strings are treated as blanks)
                var deleteOptions = new DeleteBlankOptions
                {
                    EmptyStringAsBlank = true,      // Treat cells with empty strings as blank
                    UpdateReference = true          // Update any formulas that reference the deleted rows
                };
                sheet.Cells.DeleteBlankRows(deleteOptions);

                // 8. Save the resulting workbook
                string outputPath = "SmartMarkers_IgnoringEmptyRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
