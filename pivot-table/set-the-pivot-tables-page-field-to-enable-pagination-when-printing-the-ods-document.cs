// Title: Enable pagination for ODS pivot tables with Aspose.Cells for .NET
// Description: Loads an ODS workbook, verifies the presence of a pivot table, iterates over its page fields and calls ShowReportFilterPage to create a separate report‑filter page for each field, resulting in printed pagination. The modified workbook is saved back to ODS format.
// Keywords: Aspose.Cells | ODS pivot table pagination | ShowReportFilterPage | C# pivot table page fields | print pagination ODS | Aspose.Cells .NET example | pivot table report filter page
// Common Searches: Aspose.Cells enable pagination ODS pivot | ShowReportFilterPage C# example | print each page field of a pivot table | ODS pivot table separate pages | Aspose.Cells pagination for pivot tables
// Developer Intent: Create separate printed pages for each page field of an ODS pivot table using Aspose.Cells.
// Use Cases: Generate a distinct report‑filter page per page field to improve printed layout of ODS pivot tables. | Validate worksheet content before modifying pivot tables to avoid runtime errors. | Preserve ODS formatting while adding pagination settings and saving the workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to add pagination to all page fields of a pivot table in an ODS file. | Explain the effect of ShowReportFilterPage on printing ODS pivot tables with Aspose.Cells. | Suggest robust error‑handling patterns for loading and saving ODS workbooks that contain pivot tables.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Loads an ODS workbook, verifies the presence of a pivot table, iterates over its page fields and calls ShowReportFilterPage to create a separate report‑filter page for each field, resulting in printed pagination. The modified workbook is saved back to ODS format.
    public class EnablePivotTablePagination
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.ods";
            const string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load an existing ODS workbook that contains a pivot table
                Workbook workbook = new Workbook(inputPath);

                // Assume the pivot table is on the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table
                if (sheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = sheet.PivotTables[0];

                // For each page field, generate a separate report filter page.
                // This causes pagination when the document is printed.
                foreach (PivotField pageField in pivotTable.PageFields)
                {
                    pivotTable.ShowReportFilterPage(pageField);
                }

                // Save the workbook back to ODS format
                workbook.Save(outputPath, SaveFormat.Ods);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
