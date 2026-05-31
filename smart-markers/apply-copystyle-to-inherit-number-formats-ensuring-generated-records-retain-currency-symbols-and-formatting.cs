using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyStyleDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // ---------- Source worksheet ----------
                Worksheet srcSheet = workbook.Worksheets[0];
                srcSheet.Name = "Source";

                // Put a numeric value in A1
                Cell srcCell = srcSheet.Cells["A1"];
                srcCell.PutValue(1234.56);

                // Create a style with a currency number format (built‑in format 5)
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Number = 5; // "$#,##0_);($#,##0)" format
                currencyStyle.IsNumberFormatApplied = true;

                // Apply the style to a source range (A1:B2 for demonstration)
                Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:B2");
                srcRange.SetStyle(currencyStyle);

                // ---------- Destination worksheet ----------
                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Put a numeric value in the same cell location
                Cell destCell = destSheet.Cells["A1"];
                destCell.PutValue(9876.54);

                // Create a destination range of the same size
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1:B2");

                // Copy the style (including number format) from source range to destination range
                destRange.CopyStyle(srcRange);

                // Define output file path
                string outputPath = "CopyStyleCurrencyDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}