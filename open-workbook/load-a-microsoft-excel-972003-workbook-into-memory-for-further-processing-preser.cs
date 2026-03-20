using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel 97‑2003 workbook (.xls)
            string filePath = "sample.xls";

            // Load the workbook while preserving its original formatting
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet to demonstrate successful load
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine("Loaded worksheet name: " + firstSheet.Name);

            // Additional processing can be performed on 'workbook' here
        }
    }
}