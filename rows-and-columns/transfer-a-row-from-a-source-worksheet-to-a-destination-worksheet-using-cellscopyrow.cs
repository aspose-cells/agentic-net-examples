using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRowTransfer
{
    public class TransferRowExample
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook which will contain both source and destination worksheets
                Workbook workbook = new Workbook();

                // Access the first worksheet and treat it as the source sheet
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Populate some data in the source row (row index 0)
                sourceSheet.Cells["A1"].PutValue("Item");
                sourceSheet.Cells["B1"].PutValue(123);
                sourceSheet.Cells["C1"].PutValue(DateTime.Now);

                // Add a second worksheet to act as the destination sheet
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Define source and destination row indices (zero‑based)
                int sourceRowIndex = 0;   // first row in source sheet
                int destRowIndex = 2;     // copy to third row in destination sheet

                // Use Cells.CopyRow to transfer the entire row, including data and formatting
                destSheet.Cells.CopyRow(sourceSheet.Cells, sourceRowIndex, destRowIndex);

                // Prepare output file path
                string outputPath = "RowTransferResult.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to verify the result
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
                throw;
            }
        }
    }
}