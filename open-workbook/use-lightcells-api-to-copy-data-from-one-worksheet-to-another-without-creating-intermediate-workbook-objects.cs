// Title: Copy the first worksheet from source.xlsx to dest.xlsx using Aspose.Cells Worksheet.Copy in C#
// Common Searches: how to copy first sheet from one Excel file to another using Aspose.Cells C# | Aspose.Cells copy worksheet without losing formatting example C# | C# duplicate a worksheet from source.xlsx into dest.xlsx with Aspose.Cells | using Aspose.Cells to copy a worksheet between workbooks and save changes | Aspose.Cells copy sheet when destination file may not exist C#
// Tags: worksheet copy Aspose.Cells C# | copy worksheet between Excel workbooks Aspose.Cells | ensure destination worksheet exists Aspose.Cells | load and save Excel files Aspose.Cells C# | Worksheet.Copy method example

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the source file, creates an empty destination workbook if needed, loads both workbooks, ensures the destination has a worksheet at the target index, copies the first worksheet from the source to the destination using Worksheet.Copy, and saves the updated destination file.
class LightCellsCopyExample
{
    static void Main()
    {
        try
        {
            // Paths to the source and destination Excel files
            string sourcePath = "source.xlsx";
            string destinationPath = "dest.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Ensure the destination file exists; create an empty workbook if it does not
            if (!File.Exists(destinationPath))
            {
                var emptyWb = new Workbook();
                emptyWb.Save(destinationPath);
            }

            // Load workbooks
            var sourceWb = new Workbook(sourcePath);
            var destWb = new Workbook(destinationPath);

            // Indexes of the worksheets to copy (0‑based)
            int sourceSheetIndex = 0;      // first worksheet in source file
            int destinationSheetIndex = 0; // first worksheet in destination file

            // Ensure the destination workbook has enough worksheets
            if (destWb.Worksheets.Count <= destinationSheetIndex)
                destWb.Worksheets.Add();

            // Copy the source worksheet to the destination worksheet
            destWb.Worksheets[destinationSheetIndex].Copy(sourceWb.Worksheets[sourceSheetIndex]);

            // Save the changes
            destWb.Save(destinationPath);

            Console.WriteLine("Worksheet copied successfully using standard API.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
