// Title: Save a Gantt‑Chart Workbook as a New XLSX File to a Specified Folder with Aspose.Cells for .NET (C#)
// Description: C# example that checks the source path, creates the target directory if needed, loads the workbook containing a Gantt chart with Aspose.Cells, and saves it as "GanttChart_Output.xlsx" in the chosen folder while handling errors gracefully.
// Keywords: Aspose.Cells | C# | save workbook | Gantt chart | output folder | CreateDirectory | SaveFormat.Xlsx | Excel export automation | .NET Excel library
// Common Searches: Aspose.Cells save workbook to custom folder C# | How to export a Gantt chart Excel file using Aspose.Cells | C# create directory then save Excel workbook | Save existing Excel file as new XLSX with Aspose.Cells .NET
// Developer Intent: Export an existing workbook that contains a Gantt chart to a new XLSX file in a user‑defined directory.
// Use Cases: Provide project stakeholders with a separate Gantt‑chart file stored in a shared reports folder. | Automate nightly archiving of the latest Gantt schedule for version control and backup. | Integrate the method into a CI/CD pipeline to generate a Gantt‑chart artifact after each build.
// AI Prompts: Generate a C# function that loads an Excel workbook with a Gantt chart using Aspose.Cells and saves it as a new XLSX file in a specified output directory, ensuring the folder exists and handling missing source files. | Modify the SaveGanttChart example to return the full path of the saved file and log detailed error information instead of only writing to the console. | Write unit tests for SaveGanttChart that verify folder creation, successful save, and proper handling of a non‑existent source workbook.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsGanttSaveDemo
{
    // C# example that checks the source path, creates the target directory if needed, loads the workbook containing a Gantt chart with Aspose.Cells, and saves it as "GanttChart_Output.xlsx" in the chosen folder while handling errors gracefully.
    public class GanttChartSaver
    {
        /// <param name="sourceFilePath">Full path to the source workbook that contains the Gantt chart.</param>
        /// <param name="outputFolder">Folder where the new XLSX file will be created.</param>
        public static void SaveGanttChart(string sourceFilePath, string outputFolder)
        {
            try
            {
                // Verify that the source workbook exists.
                if (!File.Exists(sourceFilePath))
                {
                    Console.WriteLine($"Source file not found: {sourceFilePath}");
                    return;
                }

                // Ensure the output directory exists.
                Directory.CreateDirectory(outputFolder);

                // Load the existing workbook.
                Workbook workbook = new Workbook(sourceFilePath);

                // Construct the full path for the new file.
                string outputFilePath = Path.Combine(outputFolder, "GanttChart_Output.xlsx");

                // Save the workbook as XLSX.
                workbook.Save(outputFilePath, SaveFormat.Xlsx);

                Console.WriteLine($"Gantt chart workbook saved successfully to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while saving the Gantt chart: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                // Path to the workbook that already contains the Gantt chart.
                string sourcePath = @"C:\Input\ProjectPlanWithGantt.xlsx";

                // Desired output folder.
                string outputDir = @"C:\Output\GanttExports";

                SaveGanttChart(sourcePath, outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
