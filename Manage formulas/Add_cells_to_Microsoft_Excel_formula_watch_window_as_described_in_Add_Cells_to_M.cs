using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the existing XLSX workbook
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook (XLSX format) using the standard constructor
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (you can iterate over all worksheets if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Add cells to the formula watch window.
            // Example: watch cell B2 and cell D5.
            // Using the Add(string) overload of CellWatchCollection.
            int watchIndexB2 = sheet.CellWatches.Add("B2");
            int watchIndexD5 = sheet.CellWatches.Add("D5");

            // Optionally retrieve the CellWatch objects to verify or modify properties
            CellWatch watchB2 = sheet.CellWatches[watchIndexB2];
            CellWatch watchD5 = sheet.CellWatches[watchIndexD5];

            // Display watch information (row/column are zero‑based)
            Console.WriteLine($"Watch added for {watchB2.CellName} at row {watchB2.Row}, column {watchB2.Column}");
            Console.WriteLine($"Watch added for {watchD5.CellName} at row {watchD5.Row}, column {watchD5.Column}");

            // Save the modified workbook to a new file
            string outputPath = "OutputWorkbook_WithWatches.xlsx";
            workbook.Save(outputPath);
        }
    }
}