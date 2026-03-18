using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DefaultTypefaceHtmlExport
    {
        public static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Set the default font for HTML export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DefaultFontName = "Arial";

            // Save the workbook as HTML
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML at '{outputPath}'.");
        }
    }
}