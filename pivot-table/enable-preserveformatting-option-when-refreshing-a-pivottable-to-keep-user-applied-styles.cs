// Title: Preserve PivotTable Formatting on Refresh with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable PivotTable.PreserveFormatting in Aspose.Cells for .NET, apply a custom style, refresh the data, and save the workbook without losing user‑defined formatting.
// Keywords: Aspose.Cells PreserveFormatting | PivotTable formatting .NET | C# Aspose.Cells pivot refresh | keep pivot style after refresh | FormatAll Aspose.Cells | Excel pivot table styling C# | Aspose.Cells PivotTable example
// Common Searches: Aspose.Cells preserve pivot formatting after refresh | Set PreserveFormatting on PivotTable using C# | Apply custom style to Aspose.Cells pivot table | Refresh pivot data without losing styles Aspose.Cells | PivotTable.FormatAll example C#
// Developer Intent: Enable the PreserveFormatting flag on a PivotTable so that any custom styles remain intact when the table is refreshed or recalculated.
// Use Cases: Create a new workbook, add sample data and a pivot table, enable PreserveFormatting, style the data area, refresh, and export to XLSX. | Open an existing Excel file containing a pivot table, set PreserveFormatting = true, refresh the source data, and save without overwriting user formatting. | Apply a reusable cell style (font, size, background) to the entire pivot table using FormatAll while ensuring the style persists across multiple data refreshes. | Automate daily report generation where pivot tables are refreshed but corporate branding (colors, fonts) must stay consistent.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, sets PivotTable.PreserveFormatting = true, applies a bold Arial style to the data area, refreshes the pivot, and saves the file. | Explain the effect of the PreserveFormatting property in Aspose.Cells PivotTables and show a complete example of preserving custom formatting after RefreshData and CalculateData. | Provide step‑by‑step instructions for creating a sample workbook, adding a pivot table, enabling PreserveFormatting, styling the pivot, and exporting the result using Aspose.Cells for .NET. | Write a GitHub‑ready snippet that demonstrates using PivotTable.FormatAll together with PreserveFormatting to keep a light‑blue background on refreshed pivot tables.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable PivotTable.PreserveFormatting in Aspose.Cells for .NET, apply a custom style, refresh the data, and save the workbook without losing user‑defined formatting.
    public class PivotTablePreserveFormattingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string sourcePath = "source.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found. Creating a sample workbook.");

                // Create a sample workbook with data and a pivot table
                var sampleWorkbook = new Workbook();
                var ws = sampleWorkbook.Worksheets[0];
                ws.Name = "Data";

                // Populate sample data
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Amount");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["B2"].PutValue(100);
                ws.Cells["A3"].PutValue("B");
                ws.Cells["B3"].PutValue(150);
                ws.Cells["A4"].PutValue("A");
                ws.Cells["B4"].PutValue(200);
                ws.Cells["A5"].PutValue("B");
                ws.Cells["B5"].PutValue(250);

                // Add a pivot table based on the sample data
                var pivotSheet = sampleWorkbook.Worksheets.Add("Pivot");
                var samplePivotTable = pivotSheet.PivotTables[pivotSheet.PivotTables.Add("=Data!$A$1:$B$5", "C3", "PivotTable1")];
                samplePivotTable.RowFields.Add(samplePivotTable.RowFields[0]); // Category as row field
                samplePivotTable.DataFields.Add(samplePivotTable.DataFields[0]); // Amount as data field

                // Save the sample source workbook
                sampleWorkbook.Save(sourcePath, SaveFormat.Xlsx);
                Console.WriteLine($"Sample workbook created at \"{sourcePath}\".");
            }

            try
            {
                // Load the workbook containing the pivot table
                var workbook = new Workbook(sourcePath);

                // Access the first worksheet (adjust index if needed)
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the first worksheet.");
                    return;
                }

                // Get the first pivot table
                var pivotTable = worksheet.PivotTables[0];

                // Enable preserving formatting when the pivot table is refreshed or recalculated
                pivotTable.PreserveFormatting = true;

                // OPTIONAL: Apply a custom style to the pivot table data area
                var style = workbook.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 10;
                style.Font.IsBold = true;
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;
                pivotTable.FormatAll(style);

                // Refresh the pivot table data from its source and recalculate
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}
