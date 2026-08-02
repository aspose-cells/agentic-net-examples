using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string excelFilePath = "input.xlsx";

            // Load the workbook into memory using the string constructor
            Workbook workbook = new Workbook(excelFilePath);

            // Access the first worksheet for further manipulation
            Worksheet firstSheet = workbook.Worksheets[0];

            // Example manipulation: read a cell value
            string cellValue = firstSheet.Cells["A1"].StringValue;
            Console.WriteLine($"Value in A1: {cellValue}");

            // The workbook is now loaded in memory and can be used for additional operations
        }
    }
}