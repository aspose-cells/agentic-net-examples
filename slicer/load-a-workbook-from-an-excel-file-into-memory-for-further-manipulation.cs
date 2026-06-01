using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string excelFilePath = "input.xlsx";

            // Load the workbook into memory using the Workbook(string) constructor
            Workbook workbook = new Workbook(excelFilePath);

            // The workbook is now in memory and can be manipulated.
            // Example: read the value of cell A1 from the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            string cellValue = firstSheet.Cells["A1"].StringValue;
            Console.WriteLine($"Value of A1: {cellValue}");

            // Additional manipulation can be performed here...

            // (Optional) Save the workbook after modifications
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}