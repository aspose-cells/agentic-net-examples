using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDebug
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a sample formula in cell A1
            worksheet.Cells["A1"].Formula = "=1+2+3";

            // Initially hide formulas (show calculated results)
            worksheet.ShowFormulas = false;
            Console.WriteLine("ShowFormulas OFF -> Cell A1 displays: " + worksheet.Cells["A1"].StringValue);

            // Enable formula view for debugging
            worksheet.ShowFormulas = true;
            Console.WriteLine("ShowFormulas ON  -> Cell A1 displays: " + worksheet.Cells["A1"].StringValue);

            // Save the workbook to verify the setting persists when opened in Excel
            workbook.Save("FormulaDebugDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}