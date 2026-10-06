// Title: Save a workbook containing a Gantt chart as a new XLSX file to a specific output folder using Aspose.Cells for .NET
// AI Prompts: Load an existing Excel workbook that includes a Gantt chart and persist it as a new .xlsx file in a designated output directory with Aspose.Cells. | Create the target output folder if it does not exist, then save the workbook using SaveFormat.Xlsx. | When the source workbook cannot be found, generate a new workbook with a worksheet named "Gantt" before saving it to the specified path.
// Common Searches: aspocells c# save workbook to custom folder with gantt chart | how to create output directory and save excel file using Aspose.Cells .net | load existing workbook or create new one if missing then save as xlsx aspocells | save excel workbook with gantt chart to specific path aspocells c# example
// Tags: Aspose.Cells save workbook to custom folder | C# create output directory for Excel file | load workbook with Gantt chart Aspose.Cells | SaveFormat.Xlsx workbook persistence | handle missing source workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// This example loads an existing Excel workbook that contains a Gantt chart (or creates a new workbook with a "Gantt" worksheet if the file is absent), ensures the target output directory exists, and saves the workbook as a new .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the existing workbook that already contains the Gantt chart
        string inputWorkbookPath = @"C:\Input\GanttChartWorkbook.xlsx";

        // Specify the output folder and the new file name
        string outputFolder = @"C:\Output";
        string outputFileName = "GanttChartWorkbook_Saved.xlsx";
        string outputPath = Path.Combine(outputFolder, outputFileName);

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        Workbook workbook = null;

        try
        {
            // Load the workbook if the file exists; otherwise create a new workbook
            if (File.Exists(inputWorkbookPath))
            {
                workbook = new Workbook(inputWorkbookPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Gantt";
            }

            // Save the workbook as a new XLSX file in the output folder
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors (e.g., permission issues, invalid format)
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Release resources if needed
            workbook?.Dispose();
        }
    }
}
