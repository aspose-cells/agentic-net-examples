// Title: Load an Excel workbook, add a new worksheet, and create a pivot table with a dynamic source range using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, creates a new sheet called 'PivotTable', and builds a pivot table at A1 using the worksheet's MaxDisplayRange as the data source. | Demonstrate how to programmatically assign the first column as a row field and the second column as a data field in an Aspose.Cells pivot table created from a dynamic source range. | Provide a full .NET example that writes the modified workbook to disk after adding the pivot table, including proper exception handling.
// Common Searches: asp.net aspose.cells create pivot table on separate worksheet from existing workbook | c# use MaxDisplayRange to define pivot table source in Aspose.Cells | how to add row and value fields to a pivot table with Aspose.Cells API | example of saving workbook after inserting pivot table using Aspose.Cells C# | determine used range for pivot table source programmatically Aspose.Cells
// Tags: Aspose.Cells dynamic pivot table source range | C# add new worksheet for pivot table Aspose.Cells | configure row and data fields in Aspose.Cells pivot | save workbook with inserted pivot Aspose.Cells | use MaxDisplayRange for pivot data source C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The example loads an existing Excel file (or creates one with sample data), identifies the used range via MaxDisplayRange, adds a new worksheet named 'PivotTable', creates a pivot table at cell A1 using that range, assigns the first column as a row field and the second as a data field, and saves the workbook with the new pivot table.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "InputData.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one with sample data
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Name = "SourceData";

                    // Sample data
                    ws.Cells["A1"].PutValue("Category");
                    ws.Cells["B1"].PutValue("Amount");
                    ws.Cells["A2"].PutValue("A");
                    ws.Cells["B2"].PutValue(100);
                    ws.Cells["A3"].PutValue("B");
                    ws.Cells["B3"].PutValue(200);
                }

                // Ensure the source worksheet is correctly named
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceData";

                // Determine the used range of the source data
                Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;
                string sourceData = $"={sourceSheet.Name}!{sourceRange.Address}";

                // Add a new worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Create the pivot table at cell A1
                int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "MyPivotTable");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields if columns are available
                if (sourceRange.ColumnCount > 0)
                    pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                if (sourceRange.ColumnCount > 1)
                    pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Save the workbook with the new pivot table
                string outputPath = "OutputWithPivot.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
