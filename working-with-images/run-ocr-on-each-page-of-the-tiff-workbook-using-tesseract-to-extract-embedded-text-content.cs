// Title: Extract text from each page of a multi‑page TIFF using Tesseract OCR and save it to an Excel workbook with Aspose.Cells in C#
// Common Searches: c# code to run Tesseract OCR on every page of a multi‑page TIFF and export the results to XLSX | how to create an Excel file from OCR output of TIFF images using Aspose.Cells in .NET | sample program that reads a TIFF into memory, processes it with TesseractEngine, and writes text to a worksheet | using Aspose.Cells to store OCR‑extracted strings in specific cells of an Excel workbook | handling missing tessdata folder when performing Tesseract OCR with Aspose.Cells in C#
// Tags: tesseract ocr multi-page tiff c# | aspocells write ocr text to xlsx | c# extract text from tiff pages | ocr result to excel worksheet aspocells | process pix image with tesseractengine | save extracted text in cell a1 using aspocells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTiffOcr
{
    // Minimal stub implementations to allow compilation without the actual Tesseract library.
    // In a real scenario, reference the Tesseract NuGet package and remove these stubs.
    internal enum EngineMode { Default }

    // The example reads a multi‑page TIFF file into a byte array, initializes a TesseractEngine with the tessdata folder, loads the image into a Pix object, runs OCR to obtain the page text, creates a new Aspose.Cells Workbook, adds a worksheet named "ExtractedText", writes the OCR result into cell A1, auto‑fits the column, and saves the workbook as an XLSX file. It also includes error handling for missing files, directories, and OCR failures.
    internal class Pix : IDisposable
    {
        public static Pix LoadFromMemory(byte[] data) => new Pix();
        public void Dispose() { }
    }

    internal class Page : IDisposable
    {
        // Returns empty text; replace with actual OCR result when using Tesseract.
        public string GetText() => string.Empty;
        public void Dispose() { }
    }

    internal class TesseractEngine : IDisposable
    {
        public TesseractEngine(string dataPath, string language, EngineMode mode) { }
        public Page Process(Pix image) => new Page();
        public void Dispose() { }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsTiffOcr <tiffFilePath> [outputExcelPath]");
                return;
            }

            string tiffPath = args[0];
            if (!File.Exists(tiffPath))
            {
                Console.WriteLine($"Error: TIFF file not found at '{tiffPath}'.");
                return;
            }

            string outputExcelPath = args.Length > 1
                ? args[1]
                : Path.Combine(Path.GetDirectoryName(tiffPath) ?? Environment.CurrentDirectory,
                               Path.GetFileNameWithoutExtension(tiffPath) + "_ExtractedText.xlsx");

            // Ensure the output directory exists.
            try
            {
                string outputDir = Path.GetDirectoryName(outputExcelPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory: {ex.Message}");
                return;
            }

            // Path to the folder that contains the Tesseract language data (tessdata).
            // It is assumed that a folder named "tessdata" exists next to the executable.
            string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
            if (!Directory.Exists(tessDataPath))
            {
                Console.WriteLine($"Warning: Tesseract data folder not found at '{tessDataPath}'. OCR will return empty text.");
            }

            try
            {
                // Read the entire TIFF file into memory.
                byte[] tiffBytes;
                using (FileStream fs = new FileStream(tiffPath, FileMode.Open, FileAccess.Read))
                {
                    tiffBytes = new byte[fs.Length];
                    fs.Read(tiffBytes, 0, tiffBytes.Length);
                }

                // Initialize Tesseract engine (English language as example).
                using (var ocrEngine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
                {
                    // Load the image into a Pix object for Tesseract.
                    using (var pix = Pix.LoadFromMemory(tiffBytes))
                    {
                        string extractedText = string.Empty;
                        try
                        {
                            using (var page = ocrEngine.Process(pix))
                            {
                                extractedText = page.GetText();
                            }
                        }
                        catch (Exception ocrEx)
                        {
                            Console.WriteLine($"OCR failed: {ocrEx.Message}");
                        }

                        Console.WriteLine("--- OCR Result ---");
                        Console.WriteLine(string.IsNullOrWhiteSpace(extractedText) ? "[No text extracted]" : extractedText.TrimEnd());

                        // Create a new workbook to store extracted text.
                        Workbook workbook = new Workbook();
                        workbook.Worksheets.Clear(); // Remove default sheet.

                        // Add a worksheet and put the extracted text into cell A1.
                        Worksheet sheet = workbook.Worksheets.Add("ExtractedText");
                        sheet.Cells["A1"].PutValue(extractedText);
                        sheet.AutoFitColumn(0);

                        // Save the workbook with extracted text.
                        workbook.Save(outputExcelPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Extraction complete. Workbook saved to '{outputExcelPath}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
