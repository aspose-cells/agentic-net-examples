using System;
using Aspose.Cells;

namespace AsposeCellsMergeAndExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example: merge cells from row 0, column 0 spanning 4 rows and 3 columns (A1:C4)
            cells.Merge(firstRow: 0, firstColumn: 0, totalRows: 4, totalColumns: 3);

            // Optionally put a value in the merged cell to demonstrate the result
            cells[0, 0].PutValue("Merged Area");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Merge contiguous empty TD elements to reduce HTML size
                MergeEmptyTdType = MergeEmptyTdType.MergeForcely,

                // Export only the active worksheet (optional)
                ExportActiveWorksheetOnly = true,

                // Ensure grid lines are exported for better visual similarity
                ExportGridLines = true
            };

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook merged and exported to HTML successfully: {outputPath}");
        }
    }
}