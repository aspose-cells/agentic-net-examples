using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    class RefreshPivotAfterConsolidationChange
    {
        static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data for the pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);

                // Add a pivot table based on the source range
                int ptIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
                PivotTable pivot = sheet.PivotTables[ptIndex];

                // Add a row field (Category) and a data field (Amount)
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Change the consolidation function of the data field (e.g., from Sum to Average)
                PivotField dataField = pivot.DataFields[0];
                dataField.Function = ConsolidationFunction.Average;

                // Refresh the pivot cache and recalculate the pivot table to reflect the new function
                pivot.RefreshData();
                pivot.CalculateData();

                // Ensure output directory exists
                string outputPath = "PivotConsolidationRefresh.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}