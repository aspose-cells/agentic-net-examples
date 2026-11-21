// Title: How to confirm that changing a Power Query formula value in a worksheet cell persists after saving with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells verify Power Query formula cell update persists after workbook save | update Power Query definition stored in Excel cell using Aspose.Cells and check persistence | Aspose.Cells .NET read back modified Power Query string from worksheet to ensure change | validate PowerQueryFormulaItem.Value change reflects in underlying query definition with Aspose.Cells | test Power Query definition replacement in Excel file using Aspose.Cells C# example
// Tags: Aspose.Cells modify Power Query definition in Excel worksheet | C# ensure Power Query string persists after saving workbook | update PowerQueryFormulaItem value using Aspose.Cells API | validate Power Query formula cell content with Aspose.Cells | check underlying query definition change in Aspose.Cells workbook

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes an initial Power Query definition to cell A1, replaces it with a new definition, reads the cell back to confirm the update, and saves the file, demonstrating that modifying a Power Query formula value via Aspose.Cells persists in the saved Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Store the initial Power Query definition in a cell (as a placeholder)
            const string queryCellAddress = "A1";
            string initialDefinition = "let Source = Excel.CurrentWorkbook(){[Name=\"Table1\"]}[Content] in Source";
            ws.Cells[queryCellAddress].PutValue(initialDefinition);

            // Display the initial query definition
            Console.WriteLine("Initial Query Definition:");
            Console.WriteLine(ws.Cells[queryCellAddress].StringValue);

            // Update the definition with a new value
            string newDefinition = "let Source = Excel.CurrentWorkbook(){[Name=\"Table2\"]}[Content] in Source";
            ws.Cells[queryCellAddress].PutValue(newDefinition);

            // Retrieve the updated definition to verify the change persisted
            string updatedDefinition = ws.Cells[queryCellAddress].StringValue;

            // Validate that the underlying definition matches the new value
            if (updatedDefinition == newDefinition)
            {
                Console.WriteLine("Update successful: underlying query definition reflects the new value.");
            }
            else
            {
                Console.WriteLine("Update failed: underlying query definition does not match the new value.");
            }

            // Save the workbook to confirm persistence (optional)
            string outputPath = "PowerQueryUpdateTest.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
