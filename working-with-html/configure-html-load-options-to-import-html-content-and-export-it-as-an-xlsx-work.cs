using System;
using Aspose.Cells;

namespace HtmlToXlsxExample
{
    class Program
    {
        static void Main()
        {
            // Input HTML file path
            string inputHtmlPath = "input.html";

            // Output XLSX file path
            string outputXlsxPath = "output.xlsx";

            // Create HtmlLoadOptions and set desired import options
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.AutoFitColsAndRows = true;      // Auto‑fit columns and rows after import
            loadOptions.LoadFormulas = true;            // Import formulas present in the HTML
            loadOptions.SupportDivTag = true;           // Preserve layout of <div> tags
            loadOptions.DeleteRedundantSpaces = true;   // Remove extra spaces introduced by <br> tags
            loadOptions.ConvertNumericData = true;      // Convert numeric strings to numbers
            loadOptions.ConvertDateTimeData = true;     // Convert date strings to DateTime values

            // Load the HTML file into a Workbook using the configured options
            Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

            // Save the workbook as an XLSX file
            workbook.Save(outputXlsxPath, SaveFormat.Xlsx);
        }
    }
}