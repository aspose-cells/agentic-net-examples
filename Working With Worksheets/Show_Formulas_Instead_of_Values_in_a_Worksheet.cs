using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            ShowFormulasDemo.Run();
        }
    }

    public class ShowFormulasDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a formula in cell A1
            worksheet.Cells["A1"].Formula = "=1+2+3";

            // Display the calculated result (ShowFormulas = false)
            worksheet.ShowFormulas = false;
            Console.WriteLine("ShowFormulas OFF: " + worksheet.Cells["A1"].StringValue);

            // Toggle to display the formula text (ShowFormulas = true)
            worksheet.ShowFormulas = true;
            Console.WriteLine("ShowFormulas ON: " + worksheet.Cells["A1"].StringValue);

            // Save the workbook
            workbook.Save("ShowFormulasDemo.xlsx");
        }
    }
}