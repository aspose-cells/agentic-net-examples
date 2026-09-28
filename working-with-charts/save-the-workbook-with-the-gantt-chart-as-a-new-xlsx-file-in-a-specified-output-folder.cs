// Title: Save a workbook with a Gantt chart to a new XLSX file in a specified output folder using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing Excel file (or create a new workbook) with Aspose.Cells, verify that the target directory exists, and save the workbook as an XLSX file to a custom path. | Detect whether the source workbook is present, create the output folder on‑the‑fly if needed, and call Workbook.Save with SaveFormat.Xlsx to write the file. | Wrap the load‑and‑save sequence in a try‑catch block to log any I/O or Aspose.Cells exceptions that may occur.
// Common Searches: Aspose.Cells C# save workbook to specific folder when source file may be missing | How to create output directory and export Excel workbook with Gantt chart using Aspose.Cells | C# Aspose.Cells save workbook as new XLSX file in custom location | Saving Excel file to a different directory with Aspose.Cells .NET API
// Tags: Aspose.Cells workbook.Save to custom directory | C# ensure output folder exists before saving Excel | Aspose.Cells handling missing source file | SaveFormat.Xlsx with Aspose.Cells | export Gantt chart workbook using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel workbook (or creates a new one if the file is absent), guarantees that the specified output folder exists, and then saves the workbook—including any Gantt chart—as a new XLSX file using Aspose.Cells for .NET, with error handling for file I/O issues.
class Program
{
    static void Main(string[] args)
    {
        // Path to the existing workbook that may contain the Gantt chart
        string inputFilePath = @"C:\InputFolder\ProjectSchedule.xlsx";

        // Output folder and file name
        string outputFolder = @"C:\OutputFolder";
        string outputFileName = "ProjectSchedule_Gantt.xlsx";
        string outputFilePath = Path.Combine(outputFolder, outputFileName);

        try
        {
            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            Workbook workbook;

            // Load the workbook if it exists; otherwise create a new one
            if (File.Exists(inputFilePath))
            {
                workbook = new Workbook(inputFilePath);
            }
            else
            {
                Console.WriteLine($"Input file not found: {inputFilePath}. Creating a new workbook.");
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Save the workbook as a new XLSX file in the specified output folder
            workbook.Save(outputFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
