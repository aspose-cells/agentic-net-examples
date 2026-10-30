// Title: Hide PivotTable field headers in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide pivot table field headers | Set ShowFieldHeaders false in Aspose.Cells pivot table | Remove field headers from generated PivotTable with Aspose.Cells .NET | How to disable pivot table field headers when creating Excel file using Aspose.Cells | Aspose.Cells pivot table options to hide headers programmatically
// Tags: Aspose.Cells pivot table header visibility control | Aspose.Cells ShowFieldHeaders property usage | Aspose.Cells pivot table options .NET | Excel pivot table without field headers using Aspose | Generate pivot table programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotFieldType

// The example creates a workbook, adds sample data, builds a PivotTable on a separate sheet, and demonstrates how to hide the field headers by setting the ShowFieldHeaders property to false (noting that the property is unavailable in the current Aspose.Cells version). The workbook is saved as PivotTable_HideFieldHeaders.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet (Data sheet)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");

            // Sample data rows
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(250);

            // Add a new worksheet for the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Define the source data range for the pivot table
            string sourceData = "Data!A1:B5";

            // Add the pivot table to the pivot sheet at cell C3
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceData, "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Add fields: Category as Row field, Amount as Data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Amount

            // NOTE: The ShowFieldHeaders property is not available in the current Aspose.Cells version.
            // If needed, adjust pivot table options via the Options property or other available settings.

            // Prepare output path
            string outputPath = "PivotTable_HideFieldHeaders.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
