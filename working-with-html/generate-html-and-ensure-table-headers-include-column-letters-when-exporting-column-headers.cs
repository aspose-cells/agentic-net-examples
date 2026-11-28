// Title: Export an Excel workbook to HTML with column letters as table headers using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to HTML and include column letters in the first row | Add a header row with column letters before saving Excel as HTML using Aspose.Cells | Generate HTML from Excel in .NET with column letter headers and embedded Base64 images
// Tags: export worksheet to HTML Aspose.Cells | insert column letter header row C# | CellsHelper.ColumnIndexToName example | HtmlSaveOptions base64 image embedding | create HTML table from Excel workbook

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The sample loads or creates an Excel workbook, inserts a top row containing column letters (A, B, C, …) using CellsHelper.ColumnIndexToName, configures HtmlSaveOptions to embed images as Base64 with UTF‑8 encoding, and saves the workbook as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputFile = "input.xlsx";
            string outputFile = "output.html";

            Workbook workbook = null;

            // Load existing workbook or create a new one with sample data
            if (File.Exists(inputFile))
            {
                try
                {
                    Console.WriteLine($"Loading workbook from \"{inputFile}\"...");
                    workbook = new Workbook(inputFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found. Creating a new workbook with sample data.");
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data (starting from row 2 to leave space for column headers)
                cells["A2"].PutValue("Product");
                cells["B2"].PutValue("Quantity");
                cells["C2"].PutValue("Price");

                cells["A3"].PutValue("Apple");
                cells["B3"].PutValue(120);
                cells["C3"].PutValue(0.5);

                cells["A4"].PutValue("Banana");
                cells["B4"].PutValue(85);
                cells["C4"].PutValue(0.3);

                cells["A5"].PutValue("Cherry");
                cells["B5"].PutValue(60);
                cells["C5"].PutValue(1.2);
            }

            // Insert a header row with column letters (A, B, C, ...) before the existing data
            try
            {
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the last used column
                int lastColumn = cells.MaxColumn;

                // Insert a new row at the top
                cells.InsertRows(0, 1);

                // Fill the new first row with column letters
                for (int col = 0; col <= lastColumn; col++)
                {
                    // Convert column index to its letter representation
                    string columnLetter = CellsHelper.ColumnIndexToName(col);
                    cells[0, col].PutValue(columnLetter);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while adding column header row: {ex.Message}");
                return;
            }

            // Configure HTML save options (no separate CSS or images)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,
                Encoding = Encoding.UTF8
            };

            // Save workbook as HTML
            try
            {
                Console.WriteLine($"Saving workbook as HTML to \"{outputFile}\"...");
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine("HTML export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during HTML export: {ex.Message}");
                return;
            }

            // Verify output file and display a short preview
            if (File.Exists(outputFile))
            {
                Console.WriteLine($"Output file size: {new FileInfo(outputFile).Length} bytes");
                Console.WriteLine("First 200 characters of the generated HTML:");
                string preview = File.ReadAllText(outputFile);
                Console.WriteLine(preview.Length > 200 ? preview.Substring(0, 200) + "..." : preview);
            }
            else
            {
                Console.WriteLine("Failed to locate the generated HTML file.");
            }
        }
    }
}
