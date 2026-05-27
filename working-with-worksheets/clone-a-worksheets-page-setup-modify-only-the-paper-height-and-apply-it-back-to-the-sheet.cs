using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPageSetupCloneDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sourceSheet = workbook.Worksheets[0];

                // (Optional) Add some data to visualize the sheet
                sourceSheet.Cells["A1"].PutValue("Original Sheet");
                sourceSheet.Cells["A2"].PutValue("Data");

                // -----------------------------------------------------------------
                // Clone the page setup of the source worksheet to a temporary worksheet
                // -----------------------------------------------------------------
                // Add a temporary worksheet that will hold the cloned settings
                int tempSheetIndex = workbook.Worksheets.Add();
                Worksheet tempSheet = workbook.Worksheets[tempSheetIndex];

                // Copy all page‑setup settings from the source sheet to the temporary sheet
                // using the Copy method with default CopyOptions
                tempSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

                // -----------------------------------------------------------------
                // Modify only the paper height on the cloned page setup
                // -----------------------------------------------------------------
                // Retrieve the current paper width (read‑only) to keep it unchanged
                double currentWidth = tempSheet.PageSetup.PaperWidth;

                // Define the new paper height (in inches). For example, set it to 11.0 inches.
                double newHeight = 11.0;

                // Apply a custom paper size with the original width and the new height
                tempSheet.PageSetup.CustomPaperSize(currentWidth, newHeight);

                // -----------------------------------------------------------------
                // Apply the modified page setup back to the original worksheet
                // -----------------------------------------------------------------
                sourceSheet.PageSetup.Copy(tempSheet.PageSetup, new CopyOptions());

                // Remove the temporary worksheet (optional, keeps the workbook clean)
                workbook.Worksheets.RemoveAt(tempSheetIndex);

                // Save the workbook to demonstrate the result
                string outputPath = "ClonedPageSetupModifiedHeight.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}