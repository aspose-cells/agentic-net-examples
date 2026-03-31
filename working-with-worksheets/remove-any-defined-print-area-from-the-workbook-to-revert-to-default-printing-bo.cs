using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemovePrintAreaDemo
    {
        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: set a print area first (to demonstrate removal)
            worksheet.PageSetup.PrintArea = "A1:C10";

            // Remove the defined print area by setting it to an empty string
            worksheet.PageSetup.PrintArea = string.Empty;

            // Save the workbook to verify that the print area has been cleared
            workbook.Save("RemovePrintAreaDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemovePrintAreaDemo.Run();
        }
    }
}