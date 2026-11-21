// Title: How to set a custom calculation error message for a PivotTable using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set custom error text for PivotTable calculation errors | DisplayErrorString property example for Aspose.Cells PivotTable | Show custom error string when PivotTable calculation fails in .NET | Configure PivotTable error handling with Aspose.Cells API | Set ErrorString for a PivotTable using Aspose.Cells C#
// Tags: Aspose.Cells set PivotTable error string | DisplayErrorString property Aspose.Cells | PivotTable custom calculation error C# | Aspose.Cells error handling for PivotTables | PivotTable error message Excel generation

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example creates a workbook, adds sample data, builds a PivotTable, enables the DisplayErrorString flag, assigns a custom error message via the ErrorString property, and saves the file as PivotTableErrorString.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add data to the first worksheet (named "Data")
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(250);

            // Add a new worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "Pivot";

            // Define source data range for the pivot table
            string sourceData = "Data!A1:B5";

            // Create the pivot table (classic behavior is default)
            int pivotTableIndex = pivotSheet.PivotTables.Add(sourceData, 2, 2, "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Configure pivot fields: Category as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column B (Amount)

            // Set a custom message to display when calculation errors occur
            pivotTable.DisplayErrorString = true;
            pivotTable.ErrorString = "Error in calculation";

            // Determine output path and ensure directory exists
            string outputPath = "PivotTableErrorString.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
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
