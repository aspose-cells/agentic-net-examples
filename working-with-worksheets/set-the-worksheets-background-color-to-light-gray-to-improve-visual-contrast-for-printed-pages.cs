using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace AsposeCellsExamples
{
    public class WorksheetBackgroundColorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Obtain the PageSetup object
                PageSetup pageSetup = worksheet.PageSetup;

                // Configure ODS page background to a solid color
                OdsPageBackground odsBackground = pageSetup.ODSPageBackground;
                odsBackground.Type = OdsPageBackgroundType.Color;   // Use solid color background
                odsBackground.Color = Color.LightGray;              // Set background color to light gray

                // Save the workbook (the background color will be applied when saved as ODS)
                string outputPath = "WorksheetWithLightGrayBackground.ods";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved with light gray background to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetBackgroundColorDemo.Run();
        }
    }
}