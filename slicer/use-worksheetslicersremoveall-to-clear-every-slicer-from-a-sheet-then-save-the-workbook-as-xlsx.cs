// Title: Remove all slicers from a worksheet and save the workbook as XLSX with Aspose.Cells for .NET
// AI Prompts: Write a C# program that opens an Excel workbook, deletes every slicer on the first worksheet using the Worksheet.Slicers collection, and saves the modified file as an XLSX document with Aspose.Cells. | Show how to iterate through a Worksheet's SlicerCollection in reverse order to safely remove all slicers before exporting the workbook. | Provide C# code that ensures the output folder exists, clears all slicers from a sheet, and writes the workbook to a new XLSX file using Aspose.Cells.
// Common Searches: Aspose.Cells C# delete all slicers from a specific worksheet | how to clear slicer collection in Aspose.Cells before saving workbook | remove slicers programmatically with Aspose.Cells .NET and export to XLSX | C# Aspose.Cells remove slicers and save workbook to new file
// Tags: worksheet slicer removal Aspose.Cells | clear slicer collection C# | save workbook as xlsx after slicer deletion | programmatic slicer deletion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for SlicerCollection

// The example loads an existing XLSX file (or creates a new workbook if none exists), ensures at least one worksheet is present, accesses the first worksheet, iterates through its SlicerCollection in reverse order to remove each slicer, creates the output directory if needed, and saves the updated workbook as output.xlsx in XLSX format using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                workbook.Worksheets.Add();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Remove all slicers from the worksheet
            SlicerCollection slicers = sheet.Slicers;
            for (int i = slicers.Count - 1; i >= 0; i--)
            {
                slicers.RemoveAt(i);
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
