using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetDisplayRightToLeftDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable right-to-left display mode for the worksheet
                worksheet.DisplayRightToLeft = true;

                // Save the workbook
                string filePath = "RightToLeftDemo.xlsx";
                workbook.Save(filePath);

                // Reload the workbook to verify the setting, if the file exists
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
                    Console.WriteLine("DisplayRightToLeft: " + loadedWorksheet.DisplayRightToLeft);
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}