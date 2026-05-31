using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

class SubtotalPdfDemo
{
    static void Main()
    {
        try
        {
            // Path to the existing Excel file
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with Italian culture settings
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                CultureInfo = new CultureInfo("it-IT")
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the data range (A1:C6)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 5, 2); // rows 0-5, columns 0-2

            // Add subtotal: group by first column, sum third column
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 2 });

            // Save the result as PDF
            string outputPdf = "output.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine("Workbook loaded with Italian culture, subtotal added, and saved as PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}