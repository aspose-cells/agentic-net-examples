using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class AutomaticPageBreaksDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the worksheet with sample data to span multiple pages vertically
                for (int row = 0; row < 200; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Configure page setup
                worksheet.PageSetup.PrintArea = "A1:E200";
                worksheet.PageSetup.FitToPagesWide = 1;
                worksheet.PageSetup.FitToPagesTall = 0; // automatic height

                // Create print options (default options are sufficient)
                ImageOrPrintOptions printOptions = new ImageOrPrintOptions();

                // Retrieve automatic page breaks based on the current content and page setup
                CellArea[] automaticPageBreaks = worksheet.GetPrintingPageBreaks(printOptions);

                // Output information about each automatic page break
                Console.WriteLine($"Total automatic page breaks: {automaticPageBreaks.Length}");
                for (int i = 0; i < automaticPageBreaks.Length; i++)
                {
                    CellArea area = automaticPageBreaks[i];
                    Console.WriteLine($"Page {i + 1}: Starts at Row {area.StartRow + 1}, Column {area.StartColumn + 1} " +
                                      $"ends at Row {area.EndRow + 1}, Column {area.EndColumn + 1}");
                }

                // Save the workbook (optional, demonstrates that page breaks are applied when printing)
                string outputPath = "AutomaticPageBreaksDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutomaticPageBreaksDemo.Run();
        }
    }
}