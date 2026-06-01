using System;
using Aspose.Cells;

namespace HtmlToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that contains CSS background colors
            string htmlFilePath = "input.html";

            // Desired path for the resulting Excel workbook
            string excelFilePath = "output.xlsx";

            // Load the HTML file into a Workbook.
            // LoadOptions with LoadFormat.Html ensures that the HTML parser is used.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // The Aspose.Cells HTML parser automatically reads CSS styles,
            // including cell background colors, and applies them to the corresponding cells.
            // No additional options are required for preserving background colors.

            // Save the workbook in XLSX format.
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file has been successfully converted to Excel: {excelFilePath}");
        }
    }
}