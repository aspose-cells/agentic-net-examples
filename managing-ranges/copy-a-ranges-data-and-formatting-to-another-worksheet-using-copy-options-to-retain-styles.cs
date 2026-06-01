using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the source worksheet and give it a name
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Add a destination worksheet
                Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                destSheet.Name = "Destination";

                // Populate the source range with data
                Cells srcCells = sourceSheet.Cells;
                srcCells["A1"].PutValue("Hello");
                srcCells["A2"].PutValue(123);
                srcCells["B1"].PutValue(DateTime.Now);
                srcCells["B2"].PutValue(456.78);

                // Create a style to apply to the source cells
                Style style = workbook.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 14;
                style.Font.IsBold = true;
                style.ForegroundColor = System.Drawing.Color.Yellow;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the source cells
                srcCells["A1"].SetStyle(style);
                srcCells["A2"].SetStyle(style);
                srcCells["B1"].SetStyle(style);
                srcCells["B2"].SetStyle(style);

                // Define source and destination ranges
                AsposeRange sourceRange = srcCells.CreateRange("A1:B2");
                AsposeRange destinationRange = destSheet.Cells.CreateRange("C3:D4");

                // Configure paste options to copy everything (data + formatting)
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All
                };

                // Perform the copy with the specified options
                destinationRange.Copy(sourceRange, pasteOptions);

                // Define output file path
                string outputPath = "RangeCopyWithStyles.xlsx";

                // Save the workbook (overwrite if exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}