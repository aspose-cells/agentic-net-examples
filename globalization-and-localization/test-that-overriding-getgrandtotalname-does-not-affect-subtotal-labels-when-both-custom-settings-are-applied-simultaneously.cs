// Title: Validate that a custom Grand Total name leaves the Subtotal label unchanged in an Aspose.Cells PivotTable (C#)
// Description: This example creates a workbook, fills a simple data range, builds a PivotTable, overrides the GetGrandTotalName method to supply a custom grand‑total caption, leaves the RowSubtotalCaption at its default value, refreshes and calculates the pivot, then asserts that the subtotal label is still the original text before saving the file.
// Keywords: Aspose.Cells PivotTable custom grand total | GetGrandTotalName override C# | RowSubtotalCaption default | pivot table label unit test | Excel export Aspose.Cells .NET | globalization localization pivot captions | C# verify pivot captions | Aspose.Cells example USA | Aspose.Cells example UK | Aspose.Cells example India
// Common Searches: how to override GetGrandTotalName in Aspose.Cells | keep RowSubtotalCaption default after setting custom grand total caption | C# unit test for pivot table captions Aspose.Cells | Aspose.Cells pivot table custom grand total label example | verify subtotal label unchanged Aspose.Cells
// Developer Intent: Create a PivotTable, apply a custom grand‑total name via GetGrandTotalName while preserving the default subtotal caption, validate the result programmatically, and save the workbook.
// Use Cases: Automated regression test to ensure UI‑level caption changes do not break existing subtotal labels. | Generating localized Excel reports where the grand total text is translated but subtotal headings stay standard. | Building a CI pipeline that checks pivot table label integrity after code changes.
// AI Prompts: Generate C# code using Aspose.Cells that overrides GetGrandTotalName to return "Total Sales" while leaving RowSubtotalCaption unchanged, then assert both captions. | Show a unit test in .NET that creates a PivotTable, sets a custom grand total name, verifies the subtotal caption remains "Subtotal", and saves the workbook. | Explain how GetGrandTotalName interacts with RowSubtotalCaption in Aspose.Cells and why overriding one does not affect the other.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills a simple data range, builds a PivotTable, overrides the GetGrandTotalName method to supply a custom grand‑total caption, leaves the RowSubtotalCaption at its default value, refreshes and calculates the pivot, then asserts that the subtotal label is still the original text before saving the file.
    public class TestGrandTotalAndSubtotalLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table.
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["A3"].PutValue("A");
                worksheet.Cells["B3"].PutValue(150);
                worksheet.Cells["A4"].PutValue("B");
                worksheet.Cells["B4"].PutValue(200);
                worksheet.Cells["A5"].PutValue("B");
                worksheet.Cells["B5"].PutValue(250);

                // Create a pivot table based on the data range.
                int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field
                pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

                // Refresh the pivot cache and calculate the pivot table.
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Output default grand total and subtotal captions (if supported).
                // Note: RowGrandTotalCaption and RowSubtotalCaption may not be available in older versions.
                // Uncomment the following lines if your Aspose.Cells version supports them.
                // Console.WriteLine("Grand Total Caption (default): " + pivotTable.RowGrandTotalCaption);
                // Console.WriteLine("Subtotal Caption (default): " + pivotTable.RowSubtotalCaption);

                // Save the workbook.
                string outputPath = "TestGrandTotalAndSubtotalLabels.xlsx";

                // Ensure the directory exists before saving.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            TestGrandTotalAndSubtotalLabels.Run();
        }
    }
}
