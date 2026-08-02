// Title: Auto‑detect Excel workbook format using Aspose.Cells Workbook(string path) in C#
// Description: Demonstrates loading any supported Excel file with only a file path, letting Aspose.Cells infer the format, retrieving the FileFormat property, and optionally saving the workbook as PDF.
// Keywords: Aspose.Cells | Workbook constructor | auto detect format | C# | load Excel file | FileFormat property | convert to PDF | xls | xlsx | csv
// Common Searches: Aspose.Cells detect workbook format from path | C# open Excel file without specifying format | Workbook(string) auto format detection | How to get FileFormat after loading workbook | Convert Excel to PDF after auto‑detecting format
// Developer Intent: Open a workbook by providing only its file path and let Aspose.Cells automatically determine the file type.
// Use Cases: Read an unknown Excel, CSV, or XLSX file and identify its exact format. | Validate successful loading by checking the detected FileFormat before processing. | Convert a workbook loaded without explicit format to PDF or another supported type.
// AI Prompts: Generate C# code that opens a workbook using Aspose.Cells with just a file path, prints the detected format, and saves it as PDF. | Explain which file extensions Aspose.Cells can auto‑detect when using the Workbook(string) constructor. | Show how to handle an exception when the constructor encounters an unsupported or corrupted file.

using System;
using Aspose.Cells;

// Demonstrates loading any supported Excel file with only a file path, letting Aspose.Cells infer the format, retrieving the FileFormat property, and optionally saving the workbook as PDF.
class DetectWorkbookFormatDemo
{
    static void Main()
    {
        // Path to the source workbook (any supported Excel format)
        string inputPath = "sample.xlsx";

        // Load the workbook by providing only the file path.
        // The constructor automatically detects the file format.
        Workbook workbook = new Workbook(inputPath);

        // Display the detected file format.
        Console.WriteLine($"Detected file format: {workbook.FileFormat}");

        // Example: save the workbook to a different format to verify loading succeeded.
        string outputPath = "converted.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
        Console.WriteLine($"Workbook saved as: {outputPath}");
    }
}
