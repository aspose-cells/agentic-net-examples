using System;
using Aspose.Cells;

namespace ExportExcelToHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data and a formula
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].Formula = "=A1+B1";

            // Configure HTML save options:
            // - CalculateFormula = true ensures formulas are evaluated before saving.
            // - ExportFormula = false makes the exported HTML contain the calculated values,
            //   not the original formula strings.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                CalculateFormula = true,
                ExportFormula = false
            };

            // Save the workbook as HTML with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Excel exported to HTML with formulas displayed as values.");
        }
    }
}