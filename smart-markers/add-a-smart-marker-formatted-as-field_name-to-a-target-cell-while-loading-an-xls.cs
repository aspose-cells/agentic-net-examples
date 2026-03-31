using System;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            // Replace "template.xlsx" with the path to your source workbook
            Workbook workbook = new Workbook("template.xlsx");

            // Access the first worksheet (or any specific worksheet as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the target cell where the smart marker will be placed
            // For example, cell B2
            Cell targetCell = worksheet.Cells["B2"];

            // Add a smart marker formatted as {{field_name}} to the target cell
            // This places the literal text "{{field_name}}" which can be used by
            // downstream processing (e.g., custom templating engines)
            targetCell.PutValue("{{field_name}}");

            // Save the modified workbook
            // Replace "output.xlsx" with the desired output file path
            workbook.Save("output.xlsx");
        }
    }
}