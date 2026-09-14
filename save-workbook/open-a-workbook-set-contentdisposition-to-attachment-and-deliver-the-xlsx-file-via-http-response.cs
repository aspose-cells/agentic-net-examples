// Title: Save an Aspose.Cells workbook as an XLSX file with OoxmlSaveOptions and ensure the output folder exists in C#
// AI Prompts: Write C# code that creates a new Aspose.Cells Workbook, puts "Hello" in cell A1 and "World" in cell B1, and saves it as an XLSX file using OoxmlSaveOptions. | Show how to check for the target directory and create it if missing before calling Workbook.Save in a C# console program. | Provide a C# example that wraps Workbook.Save in a try‑catch block to log any errors that occur while exporting an Excel file with Aspose.Cells.
// Common Searches: c# aspose.cells save workbook to xlsx with OoxmlSaveOptions example | how to create output folder before saving Aspose.Cells workbook in console app | asp.net core return excel file from Aspose.Cells with content-disposition attachment | exception handling when exporting Excel using Aspose.Cells C# | write values to specific cells using Aspose.Cells C# tutorial
// Tags: xlsx export using OoxmlSaveOptions | write cell values Aspose.Cells C# | ensure output directory exists C# | exception handling Aspose.Cells workbook save | aspose.cells console application excel generation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// // Creates a workbook, writes "Hello" to A1 and "World" to B1, guarantees the output directory exists, and saves the file as SampleOutput.xlsx with OoxmlSaveOptions while handling possible exceptions.
public class ExcelDownloadHandler
{
    // Saves a workbook with sample data to the specified file path.
    public void SendWorkbook(string outputPath)
    {
        try
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Add some sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Configure save options for Xlsx format using OoxmlSaveOptions (compatible with all versions)
            SaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook to the given file path
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            // Log or handle the exception as needed
            Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}

public class Program
{
    // Entry point required for console applications
    public static void Main(string[] args)
    {
        try
        {
            string outputFile = "SampleOutput.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(outputFile) ?? string.Empty;
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Generate and save the workbook
            ExcelDownloadHandler handler = new ExcelDownloadHandler();
            handler.SendWorkbook(outputFile);

            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
