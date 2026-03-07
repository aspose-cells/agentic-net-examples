using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossStringDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // The HtmlCrossType enumeration does not contain a 'Strikethrough' member.
            // To demonstrate setting the property, we use a valid enum value (e.g., Cross).
            // Replace with the appropriate value if a new enum member is added in the future.
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with HtmlCrossStringType set.");
        }
    }
}