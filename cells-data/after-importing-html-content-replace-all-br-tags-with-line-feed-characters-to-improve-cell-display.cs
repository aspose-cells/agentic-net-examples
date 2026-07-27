using System;
using Aspose.Cells;

namespace AsposeCellsHtmlBrReplacement
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Load the HTML file into a workbook with default options
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            // Optional: enable deletion of redundant spaces that may appear after <br> handling
            loadOptions.DeleteRedundantSpaces = true;

            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Replace all <br> tags in cell values with line feed characters
            // This improves how line breaks are displayed in Excel cells
            workbook.Replace("<br>", "\n");

            // Save the modified workbook to an Excel file
            workbook.Save("output.xlsx");
        }
    }
}