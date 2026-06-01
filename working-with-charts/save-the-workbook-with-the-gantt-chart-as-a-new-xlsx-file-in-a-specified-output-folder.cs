using System;
using System.IO;
using Aspose.Cells;

namespace GanttChartApp
{
    public class GanttChartSaver
    {
        /// <summary>
        /// Saves the provided workbook (which contains a Gantt chart) as an XLSX file
        /// in the specified output folder.
        /// </summary>
        /// <param name="workbook">The Aspose.Cells Workbook instance to save.</param>
        /// <param name="outputFolder">The folder where the file should be written.</param>
        /// <param name="fileName">The name of the output file (e.g., "GanttChart.xlsx").</param>
        public static void SaveWorkbook(Workbook workbook, string outputFolder, string fileName)
        {
            try
            {
                if (workbook == null) throw new ArgumentNullException(nameof(workbook));
                if (string.IsNullOrWhiteSpace(outputFolder)) throw new ArgumentException("Output folder is required.", nameof(outputFolder));
                if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name is required.", nameof(fileName));

                // Ensure the output directory exists.
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }

                // Combine folder and file name to get the full path.
                string fullPath = Path.Combine(outputFolder, fileName);

                // Create OoxmlSaveOptions and enable automatic directory creation.
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    CreateDirectory = true   // Creates the folder if it does not exist.
                };

                // Save the workbook in XLSX format using the specified options.
                workbook.Save(fullPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to: {fullPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths for optional template and output.
                string templatePath = "Template.xlsx";
                string outputFolder = "Output";
                string outputFileName = "GanttChart.xlsx";

                Workbook workbook;

                // Load template if it exists; otherwise create a new workbook.
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    Console.WriteLine($"Template file not found at '{templatePath}'. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // TODO: Add Gantt chart creation logic here if needed.

                // Save the workbook using the helper method.
                GanttChartSaver.SaveWorkbook(workbook, outputFolder, outputFileName);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}