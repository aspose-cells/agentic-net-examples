// Title: Auto‑Detect Excel Workbook Format Using Aspose.Cells Workbook(filePath) in C#
// Description: Shows how to load an Excel file of any supported type (XLSX, XLS, CSV, etc.) by passing only its path to the Aspose.Cells Workbook constructor. The library automatically identifies the format, accessible via Workbook.FileFormat, after confirming the file exists and reading the first worksheet.
// Keywords: Aspose.Cells | C# workbook format detection | Workbook constructor file path | auto detect Excel type | Workbook.FileFormat | load XLSX XLS CSV Aspose | detect workbook format C#
// Common Searches: Aspose.Cells detect file format from path | Load Excel file without specifying format C# | Workbook.FileFormat property example | Open unknown Excel extension Aspose.Cells | Check file exists before creating Workbook Aspose
// Developer Intent: Open an Excel workbook by providing only its file path and let Aspose.Cells infer the file format.
// Use Cases: Process incoming Excel files of unknown extensions in a batch job. | Log the detected format for auditing before further processing. | Validate file presence and safely load workbooks in automated ETL pipelines. | Read initial worksheet data after auto‑detection to confirm successful load.
// AI Prompts: Write C# code that accepts a file path, verifies the file exists, creates an Aspose.Cells Workbook, prints the detected Workbook.FileFormat, and displays the first worksheet name and cell A1 value. | Create a method that loads multiple Excel files of varying extensions using a single Workbook constructor and returns a dictionary of file names to detected formats. | Generate error‑handling logic for loading a workbook with Aspose.Cells when the file path is invalid or the format is unsupported.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to load an Excel file of any supported type (XLSX, XLS, CSV, etc.) by passing only its path to the Aspose.Cells Workbook constructor. The library automatically identifies the format, accessible via Workbook.FileFormat, after confirming the file exists and reading the first worksheet.
    public class DetectWorkbookFormatDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the Excel file (any supported format: .xlsx, .xls, .csv, etc.)
            string filePath = "sample.xlsx";

            // Verify that the file exists to prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // The Workbook constructor that takes a file path automatically detects the format.
                Workbook workbook = new Workbook(filePath);

                // After loading, the FileFormat property reflects the detected format.
                Console.WriteLine($"Loaded file: {filePath}");
                Console.WriteLine($"Detected workbook format: {workbook.FileFormat}");

                // Example: access the first worksheet to prove the workbook is usable.
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"First worksheet name: {sheet.Name}");
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
