// Title: Delete all pivot tables from an Excel worksheet with Aspose.Cells for .NET and verify removal
// Common Searches: C# code to remove pivot tables from an Excel worksheet using Aspose.Cells | Aspose.Cells delete all pivot tables and check remaining count | How to programmatically clear pivot tables in a .NET workbook | Verify no pivot tables exist after removal with Aspose.Cells | Sample code for deleting pivot tables in Aspose.Cells for .NET
// Tags: Aspose.Cells delete pivot tables | C# worksheet pivot table removal | Aspose.Cells check pivot table count | Aspose.Cells reverse loop RemoveAt | Aspose.Cells save workbook after pivot deletion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotFieldType

// The example loads an existing Excel file (or creates a sample workbook with a pivot table if missing), counts the pivot tables on the first worksheet, removes each pivot table using a reverse loop with RemoveAt, confirms that the worksheet's PivotTables collection is empty, and saves the modified workbook to a new file.
class DeletePivotTableExample
{
    static void Main()
    {
        try
        {
            string inputPath = "InputWithPivot.xlsx";
            string outputPath = "OutputWithoutPivot.xlsx";

            Workbook workbook;

            // Load existing workbook or create a sample one if the file is missing
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = CreateSampleWorkbookWithPivot();
                workbook.Save(inputPath);
                Console.WriteLine($"Input file not found. Created sample workbook '{inputPath}'.");
            }

            Worksheet worksheet = workbook.Worksheets[0];

            int initialPivotCount = worksheet.PivotTables.Count;
            Console.WriteLine($"Initial pivot tables count: {initialPivotCount}");

            // Delete all pivot tables from the worksheet
            for (int i = worksheet.PivotTables.Count - 1; i >= 0; i--)
            {
                worksheet.PivotTables.RemoveAt(i);
            }

            bool noPivotTables = worksheet.PivotTables.Count == 0;
            Console.WriteLine($"Pivot tables remaining after deletion: {worksheet.PivotTables.Count}");
            Console.WriteLine($"Deletion successful: {noPivotTables}");

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a workbook with sample data and a pivot table for demonstration purposes
    private static Workbook CreateSampleWorkbookWithPivot()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Sample data
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Amount");
        ws.Cells["A2"].PutValue("Food");
        ws.Cells["B2"].PutValue(120);
        ws.Cells["A3"].PutValue("Transport");
        ws.Cells["B3"].PutValue(80);
        ws.Cells["A4"].PutValue("Food");
        ws.Cells["B4"].PutValue(150);

        // Define source range for the pivot table
        string sourceData = "A1:B4";

        // Add pivot table at cell G1; the last parameter (true) indicates a pivot table should be created
        int pivotIndex = ws.PivotTables.Add("G1", "A1", sourceData, true);
        PivotTable pt = ws.PivotTables[pivotIndex];
        pt.Name = "PivotTable1";

        // Configure pivot fields
        pt.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pt.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        return wb;
    }
}
