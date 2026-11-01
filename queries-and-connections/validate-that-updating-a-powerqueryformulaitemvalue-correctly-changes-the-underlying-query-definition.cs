// Title: C# example: update PowerQueryFormulaItem.Value and verify QueryDefinition with Aspose.Cells
// Common Searches: how to change the Value of a PowerQueryFormulaItem in Aspose.Cells C# | verify that PowerQueryFormulaItem.Value updates the underlying query definition using Aspose.Cells | Aspose.Cells dynamic API to modify a Power Query formula item and read its QueryDefinition | C# code sample for updating Power Query formula item value in an Excel workbook | check if assigning a new Value to PowerQueryFormulaItem reflects in QueryDefinition
// Tags: Aspose.Cells PowerQueryFormulaItem value update | Power Query formula item modification .NET | dynamic access to PowerQueryFormulaItems C# | validate query definition after value change Aspose.Cells | Excel workbook Power Query update using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Microsoft.CSharp.RuntimeBinder;

// The sample creates a new Workbook, accesses the PowerQueryFormulaItems collection via dynamic, adds a Power Query formula item, displays its initial Value, assigns a new M query string to the Value property, reads the resulting QueryDefinition, compares it to the new string to confirm they match, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Use dynamic to access Power Query API (avoids compile‑time dependency)
            dynamic sheetDyn = sheet;
            dynamic queryCollection = null;
            try
            {
                queryCollection = sheetDyn.PowerQueryFormulaItems;
            }
            catch (RuntimeBinderException)
            {
                Console.WriteLine("Power Query API is not available in the current Aspose.Cells version.");
                // Save the workbook without Power Query modifications and exit
                workbook.Save("PowerQueryUpdateTest.xlsx");
                return;
            }

            // Add a Power Query formula item with an initial query definition
            dynamic queryItem = queryCollection.Add(
                "MyQuery",
                "let Source = Excel.CurrentWorkbook(){[Name=\"Table1\"]}[Content] in Source"
            );

            // Output the initial Value (should be the same as the query definition)
            Console.WriteLine("Initial Value: " + queryItem.Value);

            // Update the Value property with a new query definition
            string newQuery = "let Source = Excel.CurrentWorkbook(){[Name=\"Table2\"]}[Content] in Source";
            queryItem.Value = newQuery;

            // Retrieve the underlying query definition after the update
            string updatedDefinition = queryItem.QueryDefinition;

            // Output the updated definition
            Console.WriteLine("Updated QueryDefinition: " + updatedDefinition);

            // Validate that the update was reflected in the underlying definition
            if (updatedDefinition == newQuery)
            {
                Console.WriteLine("Success: QueryDefinition matches the new Value.");
            }
            else
            {
                Console.WriteLine("Failure: QueryDefinition does not match the new Value.");
            }

            // Save the workbook (demonstrates lifecycle handling)
            string outputPath = "PowerQueryUpdateTest.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
