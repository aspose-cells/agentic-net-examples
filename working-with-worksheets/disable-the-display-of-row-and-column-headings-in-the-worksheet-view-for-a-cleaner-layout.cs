using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableRowColumnHeadersDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Disable row and column headers visibility
                worksheet.IsRowColumnHeadersVisible = false;

                // Save the workbook
                string filePath = "WorkbookWithoutHeaders.xlsx";
                workbook.Save(filePath);

                // Verify the setting by loading the saved workbook
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];
                    Console.WriteLine("Row and Column Headers Visible: " + loadedWorksheet.IsRowColumnHeadersVisible);
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableRowColumnHeadersDemo.Run();
        }
    }
}