using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCopyDateStyleDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Ensure the template workbook exists; create a simple one if missing
                if (!File.Exists(templatePath))
                {
                    var tempWb = new Workbook();
                    var tempSheet = tempWb.Worksheets[0];
                    var tempCell = tempSheet.Cells["A1"];
                    tempCell.PutValue(DateTime.Now);
                    // Apply a date number format (e.g., "mm-dd-yyyy")
                    var style = tempCell.GetStyle();
                    style.Number = 14; // Built‑in date format
                    tempCell.SetStyle(style);
                    tempWb.Save(templatePath);
                }

                // Load the template workbook that contains the desired date format
                var templateWorkbook = new Workbook(templatePath);
                var templateSheet = templateWorkbook.Worksheets[0];

                // Define the source range that holds the date style (cell A1)
                AsposeRange sourceDateRange = templateSheet.Cells.CreateRange("A1");

                // Create a new workbook where dates will be generated
                var resultWorkbook = new Workbook();
                var resultSheet = resultWorkbook.Worksheets[0];

                // Put a date value into the destination cell (B2)
                Cell destinationCell = resultSheet.Cells["B2"];
                destinationCell.PutValue(DateTime.Now);

                // Define the destination range that corresponds to the cell with the new date
                AsposeRange destinationRange = resultSheet.Cells.CreateRange("B2");

                // Copy the date style from the template range to the destination range
                destinationRange.CopyStyle(sourceDateRange);

                // Save the resulting workbook
                resultWorkbook.Save(resultPath);
                Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}