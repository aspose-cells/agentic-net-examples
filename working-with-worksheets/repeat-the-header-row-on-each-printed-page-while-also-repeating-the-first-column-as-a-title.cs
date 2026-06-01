using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RepeatHeaderAndFirstColumnDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Title");
                worksheet.Cells["B1"].PutValue("Header 1");
                worksheet.Cells["C1"].PutValue("Header 2");
                worksheet.Cells["D1"].PutValue("Header 3");

                // Fill rows with data
                for (int row = 2; row <= 100; row++)
                {
                    // First column acts as a title for each row
                    worksheet.Cells[row, 0].PutValue("Row Title " + (row - 1));

                    // Other columns with sample data
                    for (int col = 1; col <= 3; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row - 1}C{col}");
                    }
                }

                // Configure page setup for printing
                PageSetup pageSetup = worksheet.PageSetup;
                pageSetup.PrintTitleRows = "$1:$1";      // repeat header row
                pageSetup.PrintTitleColumns = "$A:$A";   // repeat first column
                pageSetup.PrintArea = "A1:D100";         // define print area

                string outputPath = "RepeatHeaderAndFirstColumnDemo.xlsx";

                // Ensure any existing file is removed before saving
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            RepeatHeaderAndFirstColumnDemo.Run();
        }
    }
}