// Title: C# example to render a workbook to a TIFF file with Aspose.Cells and enforce a maximum file size
// Common Searches: how to render an Excel workbook to a TIFF file with Aspose.Cells and limit the output size in C# | C# code to check TIFF file size after using Aspose.Cells WorkbookRender | Aspose.Cells image rendering options to keep generated TIFF under a specific byte limit | validate generated TIFF size against a 5 MB threshold using Aspose.Cells in .NET | prevent oversized TIFF output when converting Excel to image with Aspose.Cells C#
// Tags: Aspose.Cells workbook to TIFF rendering | ImageOrPrintOptions TIFF compression LZW | validate TIFF file size C# | maximum TIFF output size Aspose.Cells | WorkbookRender size limit handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace TiffSizeValidator
{
    // The sample creates a workbook, configures ImageOrPrintOptions for 300 dpi LZW‑compressed 24‑bit TIFF output, renders the workbook to a TIFF file using WorkbookRender, and then checks the file size against a 5 MB limit, throwing an exception if the size exceeds the allowed maximum.
    class Program
    {
        // Define the maximum allowed TIFF file size (in bytes). Adjust as needed.
        private const long MaxTiffFileSizeBytes = 5 * 1024 * 1024; // 5 MB

        static void Main(string[] args)
        {
            // Paths for the generated TIFF file.
            string outputTiffPath = Path.Combine(Directory.GetCurrentDirectory(), "GeneratedOutput.tiff");

            try
            {
                // 1. Create a sample workbook with some data.
                Workbook workbook = CreateSampleWorkbook();

                // 2. Configure image rendering options for TIFF output.
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Tiff,
                    // Use LZW compression to keep file size reasonable.
                    TiffCompression = TiffCompression.CompressionLZW,
                    // Set a 24‑bit color depth for better visual quality.
                    TiffColorDepth = ColorDepth.Format24bpp,
                    // High resolution can increase size; adjust if needed.
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    // Ensure each sheet is rendered on a single page (optional).
                    OnePagePerSheet = true
                };

                // 3. Render the entire workbook to a TIFF file.
                // WorkbookRender renders all worksheets; alternatively, SheetRender can be used per sheet.
                WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
                renderer.ToImage(outputTiffPath);

                // 4. Validate the generated file size.
                ValidateTiffFileSize(outputTiffPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <returns>A Workbook instance with sample data.</returns>
        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate header row.
            ws.Cells["A1"].PutValue("Product");
            ws.Cells["B1"].PutValue("Quantity");
            ws.Cells["C1"].PutValue("Unit Price");
            ws.Cells["D1"].PutValue("Total");

            // Sample data rows.
            string[,] data = new string[,]
            {
                { "Apple", "120", "0.50" },
                { "Banana", "85", "0.30" },
                { "Orange", "150", "0.45" },
                { "Grapes", "60", "1.20" },
                { "Mango", "40", "1.75" }
            };

            for (int i = 0; i < data.GetLength(0); i++)
            {
                int row = i + 2; // Data starts at row 2.
                ws.Cells[$"A{row}"].PutValue(data[i, 0]);
                ws.Cells[$"B{row}"].PutValue(int.Parse(data[i, 1]));
                ws.Cells[$"C{row}"].PutValue(double.Parse(data[i, 2]));
                // Formula for total = Quantity * Unit Price
                ws.Cells[$"D{row}"].Formula = $"=B{row}*C{row}";
            }

            // Add a grand total row.
            int totalRow = data.GetLength(0) + 3;
            ws.Cells[$"C{totalRow}"].PutValue("Grand Total:");
            ws.Cells[$"D{totalRow}"].Formula = $"=SUM(D2:D{totalRow - 1})";

            // Calculate formulas so that the rendered image contains evaluated values.
            wb.CalculateFormula();

            // Optional: Adjust column widths for better appearance.
            ws.AutoFitColumns();

            return wb;
        }

        /// <summary>
        /// Checks the size of the generated TIFF file against the predefined limit.
        /// </summary>
        /// <param name="filePath">Full path to the TIFF file.</param>
        private static void ValidateTiffFileSize(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine("TIFF file was not created.");
                return;
            }

            FileInfo info = new FileInfo(filePath);
            long fileSize = info.Length;

            Console.WriteLine($"Generated TIFF file: {filePath}");
            Console.WriteLine($"File size: {fileSize} bytes ({fileSize / 1024.0 / 1024.0:F2} MB)");

            if (fileSize > MaxTiffFileSizeBytes)
            {
                Console.WriteLine($"WARNING: TIFF file exceeds the maximum allowed size of {MaxTiffFileSizeBytes} bytes.");
                // Depending on requirements, you could delete the file, throw an exception, or take other actions.
                // For demonstration, we will throw an exception.
                throw new InvalidOperationException($"TIFF file size ({fileSize} bytes) exceeds the limit of {MaxTiffFileSizeBytes} bytes.");
            }
            else
            {
                Console.WriteLine("TIFF file size is within the acceptable limit.");
            }
        }
    }
}
