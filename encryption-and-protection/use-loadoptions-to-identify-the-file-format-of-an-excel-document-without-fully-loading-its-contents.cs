// Title: Detect Excel format and encryption with LoadOptions – Aspose.Cells for .NET
// Description: Shows how to open an Excel file as a read‑only stream, use FileFormatUtil.DetectFileFormat to obtain FileFormatInfo (type, encryption flag, LoadFormat), reset the stream, create LoadOptions based on the detected format (e.g., disable formula parsing), and then load the workbook efficiently.
// Keywords: Aspose.Cells detect file format | LoadOptions Excel type detection | FileFormatUtil DetectFileFormat example | check Excel encryption Aspose | skip formula parsing Aspose.Cells | identify workbook format without loading
// Common Searches: detect Excel file type without opening workbook Aspose.Cells | how to check if Excel file is encrypted before loading | use LoadOptions after format detection Aspose | FileFormatUtil DetectFileFormat C# example | optimize workbook loading by disabling formula parsing
// Developer Intent: Find the exact Excel format and encryption status without fully loading the file, then configure LoadOptions for a fast subsequent load.
// Use Cases: Determine whether an unknown file is .xlsx, .xls, .csv, etc., to select the correct LoadFormat. | Read the IsEncrypted flag to decide if a password prompt is needed before opening. | Improve performance for large workbooks by creating LoadOptions (e.g., ParsingFormulaOnOpen = false) after format detection.
// AI Prompts: Generate C# code that uses Aspose.Cells to detect an Excel file's format and encryption status via FileFormatUtil, then builds LoadOptions based on the detected LoadFormat. | Explain why the stream must be rewound after DetectFileFormat before constructing a Workbook with LoadOptions. | Recommend additional LoadOptions settings that reduce memory usage and speed up loading of big Excel files after format detection.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to open an Excel file as a read‑only stream, use FileFormatUtil.DetectFileFormat to obtain FileFormatInfo (type, encryption flag, LoadFormat), reset the stream, create LoadOptions based on the detected format (e.g., disable formula parsing), and then load the workbook efficiently.
    public class DetectFormatWithLoadOptionsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the Excel file whose format we want to detect
            string filePath = "sample.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Open the file as a read‑only stream
            using (FileStream stream = File.OpenRead(filePath))
            {
                // Detect the file format without loading the workbook into memory
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(stream);
                Console.WriteLine($"Detected FileFormatType: {formatInfo.FileFormatType}");
                Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");
                Console.WriteLine($"Detected LoadFormat: {formatInfo.LoadFormat}");

                // Reset the stream position if we later need to load the workbook
                stream.Seek(0, SeekOrigin.Begin);

                // Create LoadOptions based on the detected LoadFormat
                LoadOptions loadOptions = new LoadOptions(formatInfo.LoadFormat)
                {
                    // Example: skip formula parsing to speed up loading (optional)
                    ParsingFormulaOnOpen = false
                };

                // Load the workbook using the specific LoadOptions (demonstrates usage)
                Workbook workbook = new Workbook(stream, loadOptions);
                Console.WriteLine($"Workbook loaded successfully. Sheet count: {workbook.Worksheets.Count}");
            }
        }
    }
}
