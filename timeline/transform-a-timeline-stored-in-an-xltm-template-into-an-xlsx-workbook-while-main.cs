using System;
using System.IO;
using Aspose.Cells;

namespace TimelineConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLTM template that contains a Timeline control
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TemplateWithTimeline.xltm");

            Workbook workbook;

            if (File.Exists(templatePath))
            {
                // Load the XLTM file into a Workbook instance
                workbook = new Workbook(templatePath);
            }
            else
            {
                // If the template file is missing, create a new workbook as a fallback
                Console.WriteLine($"Template file not found at '{templatePath}'. Creating a new workbook instead.");
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Verify that the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The loaded workbook has no worksheets.");
                return;
            }

            // Access the first worksheet and check for timelines
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Timelines.Count > 0)
            {
                Console.WriteLine($"Found {sheet.Timelines.Count} timeline(s) in the worksheet.");
                // Example: change the caption of the first timeline
                // sheet.Timelines[0].Caption = "Converted Timeline";
            }
            else
            {
                Console.WriteLine("No timelines were found in the worksheet.");
            }

            // Define the output XLSX file path
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertedWorkbook.xlsx");

            // Save the workbook as an XLSX file, preserving all objects including the Timeline
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Output saved to: {outputPath}");
        }
    }
}