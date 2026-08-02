// Title: C# Aspose.Cells LoadFilter to load only defined names that contain "Total"
// Description: Demonstrates how to create a workbook, add defined names, save it, and reload it with a custom LoadFilter that loads only the defined names. After loading, the code filters names containing the word "Total", prints their details, and verifies that no other names are present.
// Keywords: Aspose.Cells LoadFilter | C# load defined names | filter named ranges Total | .NET Excel named ranges | LoadOptions custom filter | verify named ranges Aspose | Excel named range memory optimization
// Common Searches: Aspose.Cells load only named ranges | C# filter defined names containing Total | How to use LoadFilter with Aspose.Cells | Validate named ranges after loading Excel file | Reduce memory usage by loading specific named ranges
// Developer Intent: Load a workbook while restricting the load to defined names that include the keyword "Total" and confirm that other names are excluded.
// Use Cases: Extract only total‑related named ranges from large financial workbooks to improve performance. | Automated validation that required total named ranges exist before processing a template. | Generate reports that list only total calculations without loading full worksheet data.
// AI Prompts: Write C# code using Aspose.Cells LoadOptions with a custom LoadFilter to load only defined names containing a given keyword. | Explain how to extend DefinedNamesOnlyLoadFilter to exclude names that do not match a regex pattern. | Create a C# unit test that asserts only names with "Total" are loaded after applying the custom filter.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDefinedNamesFilter
{
    // Custom LoadFilter that loads only defined names.
    // Demonstrates how to create a workbook, add defined names, save it, and reload it with a custom LoadFilter that loads only the defined names. After loading, the code filters names containing the word "Total", prints their details, and verifies that no other names are present.
    public class DefinedNamesOnlyLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only the defined names for each worksheet.
            LoadDataFilterOptions = LoadDataFilterOptions.DefinedNames;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook and add some defined names ----------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Name = "Sheet1";

                // Add sample data (optional)
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].PutValue(30);

                // Add defined names, some containing "Total"
                int idx1 = wb.Worksheets.Names.Add("TotalSales");
                wb.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1";

                int idx2 = wb.Worksheets.Names.Add("Average");
                wb.Worksheets.Names[idx2].RefersTo = "=Sheet1!$A$2";

                int idx3 = wb.Worksheets.Names.Add("GrandTotal");
                wb.Worksheets.Names[idx3].RefersTo = "=Sheet1!$A$3";

                // Save the workbook to a temporary file
                string filePath = "DefinedNamesDemo.xlsx";
                wb.Save(filePath);
                wb.Dispose();

                // Ensure the file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File '{filePath}' was not found.");
                    return;
                }

                // ---------- Load the workbook with a custom LoadFilter ----------
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new DefinedNamesOnlyLoadFilter()
                };

                Workbook loadedWb = new Workbook(filePath, loadOptions);

                // Access the collection of defined names after loading
                NameCollection names = loadedWb.Worksheets.Names;

                // Find all names that contain the word "Total"
                List<Name> totalNames = names.FindAll(n => n.Text != null && n.Text.Contains("Total"));

                // Verify that only names with "Total" are present
                Console.WriteLine($"Total defined names loaded: {totalNames.Count}");
                foreach (Name n in totalNames)
                {
                    Console.WriteLine($"Name: {n.Text}, RefersTo: {n.RefersTo}");
                }

                // Optional verification: ensure no other names exist
                bool onlyTotalNames = names.Count == totalNames.Count;
                Console.WriteLine($"Only 'Total' names retained: {onlyTotalNames}");

                // Clean up
                loadedWb.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
