using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HtmlToExcelConverter
    {
        public static void Main()
        {
            // Create HTML load options
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                SupportDivTag = true,
                LoadFormulas = true,
                DeleteRedundantSpaces = true
            };

            // Load the HTML file into a workbook using the specified options
            Workbook workbook = new Workbook("input.html", loadOptions);

            // Save the workbook as an XLSX file
            workbook.Save("output.xlsx");
        }
    }
}