using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotReportGenerator
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Ensure the workbook contains at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook contains no worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the first worksheet.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];

                // Create a report filter page for each page field, if any exist
                if (pivotTable.PageFields.Count > 0)
                {
                    foreach (PivotField pageField in pivotTable.PageFields)
                    {
                        pivotTable.ShowReportFilterPage(pageField);
                    }
                }
                else
                {
                    Console.WriteLine("The pivot table has no page fields.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}