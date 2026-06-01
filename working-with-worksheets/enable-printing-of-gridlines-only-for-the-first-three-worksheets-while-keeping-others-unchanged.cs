using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PrintGridlinesFirstThreeSheets
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook with default one worksheet
            Workbook workbook = new Workbook();

            // Add additional worksheets so we have at least three
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");
            workbook.Worksheets.Add("Sheet4"); // extra sheet to demonstrate unchanged behavior

            // Enable printing of gridlines for the first three worksheets only
            for (int i = 0; i < workbook.Worksheets.Count && i < 3; i++)
            {
                Worksheet ws = workbook.Worksheets[i];
                // This property determines whether gridlines are printed on the page
                ws.PageSetup.PrintGridlines = true;
            }

            string outputPath = "PrintGridlinesFirstThreeSheets.xlsx";

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                throw;
            }
        }
    }
}