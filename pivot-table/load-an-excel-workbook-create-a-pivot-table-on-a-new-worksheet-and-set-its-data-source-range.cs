// Title: Create a Pivot Table on a New Worksheet and Define Its Source Data Range with Aspose.Cells for .NET (C#)
// Common Searches: how to programmatically add a pivot table to a new worksheet using Aspose.Cells C# | setting the source data address for a pivot table in Aspose.Cells .NET | Aspose.Cells example for creating a pivot table from an existing sheet range | C# code to generate a pivot table on a separate sheet with custom data range | Aspose.Cells pivot table creation with dynamic source range
// Tags: create pivot table Aspose.Cells C# | define pivot source range Aspose.Cells | add worksheet for pivot Aspose.Cells | pivot table data address C# | Aspose.Cells pivot table from existing sheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The sample loads an existing Excel workbook, adds a new worksheet named "PivotSheet", builds a source data address from the first worksheet using specified row and column counts, creates a pivot table called "MyPivot" at cell A1 on the new sheet with that address, optionally configures fields, ensures the output directory exists, and saves the modified workbook to a new file while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Add a worksheet for the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add();
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotSheet";

            // Define source data range (adjust as needed)
            Worksheet dataSheet = workbook.Worksheets[0]; // assume first sheet holds data
            int firstRow = 0;          // A1 row index
            int firstColumn = 0;       // A1 column index
            int totalRows = 100;       // number of rows
            int totalColumns = 4;      // number of columns

            string startCell = CellsHelper.CellIndexToName(firstRow, firstColumn);
            string endCell = CellsHelper.CellIndexToName(firstRow + totalRows - 1, firstColumn + totalColumns - 1);
            string sourceData = $"{dataSheet.Name}!{startCell}:{endCell}";

            // Create pivot table at A1 on the new sheet
            int pivotTableIndex = pivotSheet.PivotTables.Add("MyPivot", "A1", sourceData);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];

            // Optional: configure pivot fields
            // pivotTable.RowFields.Add(0);
            // pivotTable.ColumnFields.Add(1);
            // pivotTable.DataFields.Add(2);

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
