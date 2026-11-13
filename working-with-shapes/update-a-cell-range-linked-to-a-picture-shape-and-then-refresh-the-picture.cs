// Title: How to link an Excel picture shape to a cell range and refresh it with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# link picture to cell range A1:B2 | refresh linked picture shape after updating cells using Aspose.Cells | set IsLinkedToCell and LinkedCell properties for a picture in Aspose.Cells | dynamic binding picture properties when Aspose.Cells version lacks linking API | programmatically update picture linked to a range with Aspose.Cells .NET
// Tags: picture linked cell range Aspose.Cells .NET | refresh picture shape Aspose.Cells C# | IsLinkedToCell API Aspose.Cells | dynamic picture binding Aspose.Cells | populate range and update picture Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads or creates an Excel workbook, accesses the first worksheet and its first picture, and uses dynamic binding to enable linking, assign the range "A1:B2" to the picture, and call Refresh. It then fills the specified cells with sample data, saves the workbook, and gracefully handles scenarios where the linking properties are unavailable in older Aspose.Cells versions.
class UpdateLinkedPicture
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Load an existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. A new workbook will be created.");
                workbook = new Workbook();
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one picture in the worksheet.
            if (sheet.Pictures.Count == 0)
            {
                Console.WriteLine("No pictures found in the worksheet. Operation aborted.");
                return;
            }

            // Retrieve the first picture.
            Picture picture = sheet.Pictures[0];

            // Attempt to link the picture to a cell range using dynamic binding.
            // This works with newer Aspose.Cells versions that expose IsLinkedToCell, LinkedCell and Refresh.
            // If the members are unavailable, the exception is caught and the program continues.
            try
            {
                dynamic dynPic = picture;
                dynPic.IsLinkedToCell = true;
                dynPic.LinkedCell = "A1:B2";
                dynPic.Refresh();
                Console.WriteLine("Picture linked to range A1:B2 and refreshed.");
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                Console.WriteLine("Linking properties are not available in the current Aspose.Cells version. Skipping linking.");
            }

            // Fill the linked (or target) range with sample data.
            for (int row = 0; row < 2; row++)          // rows 0‑1 (A and B)
            {
                for (int col = 0; col < 2; col++)      // columns 0‑1 (1 and 2)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
