using System;
using Aspose.Cells;

namespace AsposeCellsMultiLineTest
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define multi‑line text containing leading, trailing and extra internal spaces
            string multiLineText = "  First line with leading spaces  \nSecond   line   with   extra   spaces\n   Third line with trailing spaces   ";

            // Place the text into cell A1
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue(multiLineText);

            // Enable text wrapping so line breaks are displayed within the cell
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Auto‑fit the row height to show all lines
            worksheet.AutoFitRow(0);

            // Save the workbook
            string filePath = "MultiLineSpaceTest.xlsx";
            workbook.Save(filePath);

            // Reload the workbook to verify that the text and spaces are preserved
            Workbook loadedWorkbook = new Workbook(filePath);
            Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];
            string loadedText = loadedCell.StringValue;

            // Output the loaded text to the console for inspection
            Console.WriteLine("Loaded cell text:");
            Console.WriteLine("---");
            Console.WriteLine(loadedText);
            Console.WriteLine("---");
        }
    }
}