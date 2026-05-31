using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Remove all conditional formatting from each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // The ConditionalFormattings collection holds all conditional formatting rules
                // Clear the collection to prevent any conditional formatting from being exported
                sheet.ConditionalFormattings.Clear();
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all data (including values, not just the table)
                ExportDataOptions = HtmlExportDataOptions.All,

                // Optional: keep CSS styles (default) for a clean visual representation
                // DisableCss = false; // default value, shown for clarity
            };

            // Save the workbook as an HTML file without conditional formatting
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}