using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PrintAreaClearDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data (10 rows x 5 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define a print area (B2:D6)
            worksheet.PageSetup.PrintArea = "B2:D6";

            // Save the workbook as HTML exporting ONLY the defined print area
            HtmlSaveOptions optionsPrintAreaOnly = new HtmlSaveOptions();
            optionsPrintAreaOnly.ExportPrintAreaOnly = true;   // only the print area will be exported
            workbook.Save("PrintAreaOnly.html", optionsPrintAreaOnly);
            Console.WriteLine("Saved HTML with only the defined print area.");

            // -----------------------------------------------------------------
            // Clear the print area by setting it to an empty string
            // After this, the worksheet has no explicit print area.
            // When ExportPrintAreaOnly is true, the whole sheet will be exported.
            // -----------------------------------------------------------------
            worksheet.PageSetup.PrintArea = string.Empty;

            // Save the workbook again with the same ExportPrintAreaOnly option
            HtmlSaveOptions optionsAfterClear = new HtmlSaveOptions();
            optionsAfterClear.ExportPrintAreaOnly = true;   // now the whole sheet is exported
            workbook.Save("AfterClearPrintArea.html", optionsAfterClear);
            Console.WriteLine("Saved HTML after clearing the print area (entire sheet exported).");
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            PrintAreaClearDemo.Run();
        }
    }
}