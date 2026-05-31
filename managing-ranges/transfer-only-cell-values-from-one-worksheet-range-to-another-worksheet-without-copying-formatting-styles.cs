using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class TransferValuesOnly
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Source worksheet
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate source range with values (formatting will not be copied)
            srcSheet.Cells["A1"].PutValue("Sample Text");
            var style = srcSheet.Cells["A1"].GetStyle();
            style.Font.IsBold = true;
            srcSheet.Cells["A1"].SetStyle(style);
            srcSheet.Cells["B1"].PutValue(123);
            srcSheet.Cells["A2"].PutValue(DateTime.Now);
            srcSheet.Cells["B2"].PutValue(45.67);

            // Destination worksheet
            int destIndex = workbook.Worksheets.Add();
            Worksheet destSheet = workbook.Worksheets[destIndex];
            destSheet.Name = "Destination";

            // Define source and destination ranges
            AsposeRange srcRange = srcSheet.Cells.CreateRange("A1:B2");
            AsposeRange destRange = destSheet.Cells.CreateRange("C3:D4");

            // Copy only the cell values (no formatting) from source to destination
            destRange.CopyValue(srcRange);

            // Save the workbook
            string outputPath = "TransferValuesOnly.xlsx";
            workbook.Save(outputPath);
        }
    }
}