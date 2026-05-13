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

            // Merge a range of cells across rows (e.g., A1:C3)
            // Parameters: firstRow (0‑based), firstColumn (0‑based), totalRows (1‑based), totalColumns (1‑based)
            cells.Merge(firstRow: 0, firstColumn: 0, totalRows: 3, totalColumns: 3);

            // Optionally, set a value in the merged cell to demonstrate the result
            cells[0, 0].PutValue("Merged Area");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Merge contiguous empty TD elements to reduce HTML size
                MergeEmptyTdType = MergeEmptyTdType.MergeForcely,

                // Export only the active worksheet as a table
                ExportActiveWorksheetOnly = true,

                // Ensure grid lines are exported for better visual fidelity
                ExportGridLines = true
            };

            // Save the workbook as an HTML file with the specified options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook merged and exported to HTML successfully: {outputPath}");
        }
    }
}