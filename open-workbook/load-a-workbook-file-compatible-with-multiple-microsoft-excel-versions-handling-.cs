using System;
using System.IO;
using Aspose.Cells;

class WorkbookLoader
{
    // Loads a workbook, handling various Excel formats programmatically.
    public static void LoadWorkbook(string filePath)
    {
        // Detect the file format using Aspose.Cells utility.
        FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

        // Retrieve the detected LoadFormat (e.g., Xlsx, Excel97To2003, Csv, etc.).
        LoadFormat detectedLoadFormat = formatInfo.LoadFormat;

        // Create LoadOptions based on the detected format.
        LoadOptions loadOptions = new LoadOptions(detectedLoadFormat);

        // Load the workbook with the appropriate options.
        Workbook workbook = new Workbook(filePath, loadOptions);

        // Example processing: display basic information.
        Console.WriteLine($"File: {Path.GetFileName(filePath)}");
        Console.WriteLine($"Detected LoadFormat: {detectedLoadFormat}");
        Console.WriteLine($"Worksheets count: {workbook.Worksheets.Count}");

        // Save the workbook in a unified format (Xlsx) for consistency.
        string outputPath = Path.ChangeExtension(filePath, ".converted.xlsx");
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Saved converted workbook to: {outputPath}");
        Console.WriteLine();
    }

    static void Main()
    {
        // List of sample files in different Excel-compatible formats.
        string[] sampleFiles = {
            "sample.xlsx",
            "sample.xls",
            "sample.csv",
            "sample.mht"
        };

        foreach (string file in sampleFiles)
        {
            if (File.Exists(file))
            {
                LoadWorkbook(file);
            }
            else
            {
                Console.WriteLine($"File not found: {file}");
            }
        }
    }
}