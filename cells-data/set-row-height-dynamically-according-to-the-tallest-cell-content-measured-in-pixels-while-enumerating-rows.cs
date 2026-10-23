// Title: Set Excel row height in pixels dynamically from the tallest cell content with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit row height and retrieve pixel value | Convert Aspose.Cells row height from points to pixels in .NET | Enforce a minimum pixel row height after AutoFitRow using Aspose.Cells | Iterate worksheet rows and adjust Excel row height based on cell content with Aspose.Cells | How to set dynamic row height in pixels when processing an Excel file in C#
// Tags: Aspose.Cells AutoFitRow pixel conversion | row height points‑pixel mapping Aspose.Cells | minimum pixel row height enforcement Aspose.Cells | worksheet row iteration height adjustment Aspose.Cells | dynamic Excel row height calculation .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates an Excel workbook, iterates through all used rows, auto‑fits each row, reads the calculated height in points, converts it to pixels, applies a minimum pixel height if needed, converts back to points, sets the adjusted height, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Load an existing workbook if the file exists; otherwise create a new one
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a new workbook with a default worksheet
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enumerate all used rows in the worksheet
                int firstRow = worksheet.Cells.MinRow;
                int lastRow = worksheet.Cells.MaxRow;

                for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
                {
                    // Auto‑fit the current row based on the tallest cell content.
                    // Note: Aspose.Cells' AutoFitRow method accepts only the row index.
                    worksheet.AutoFitRow(rowIndex);

                    // Retrieve the height that Aspose.Cells calculated (in points)
                    double heightInPoints = worksheet.Cells.Rows[rowIndex].Height;

                    // Convert points to pixels (1 point = 1/72 inch, 96 DPI => 1 point = 96/72 pixels)
                    double heightInPixels = heightInPoints * 96.0 / 72.0;

                    // Enforce a minimum pixel height if needed
                    const double minPixels = 20.0;
                    if (heightInPixels < minPixels)
                    {
                        heightInPixels = minPixels;
                        // Convert back to points before setting the row height
                        heightInPoints = heightInPixels * 72.0 / 96.0;
                        worksheet.Cells.Rows[rowIndex].Height = heightInPoints;
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
