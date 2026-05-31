using System;
using Aspose.Cells;

namespace AsposeCellsQuotePrefixDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a cell and set a value that starts with a single quote
            Cell cell = worksheet.Cells["B10"];
            cell.PutValue("'12345"); // The leading single quote is a quote prefix

            // Enable the QuotePrefix style for the cell
            Style style = cell.GetStyle();
            style.QuotePrefix = true;
            cell.SetStyle(style);

            // Save the workbook (lifecycle save) – optional for demonstration
            workbook.Save("QuotePrefixDemo.xlsx");

            // Load the workbook back (lifecycle load) to verify the property persists
            Workbook loadedWorkbook = new Workbook("QuotePrefixDemo.xlsx");
            Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["B10"];

            // Read the QuotePrefix property from the cell's style
            bool isQuotePrefixSet = loadedCell.GetStyle().QuotePrefix;

            // Log the value to the console
            Console.WriteLine("QuotePrefix is set: " + isQuotePrefixSet);
        }
    }
}