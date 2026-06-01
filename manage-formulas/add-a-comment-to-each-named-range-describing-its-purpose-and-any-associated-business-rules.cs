using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNamedRangeComments
{
    public class NamedRangeCommentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Define named range "SalesData"
                int salesIndex = workbook.Worksheets.Names.Add("SalesData");
                Name salesRange = workbook.Worksheets.Names[salesIndex];
                salesRange.RefersTo = "=Data!$B$2:$M$2";
                salesRange.Comment = "Monthly sales figures (numeric, >=0)";

                // Define named range "RegionList"
                int regionIndex = workbook.Worksheets.Names.Add("RegionList");
                Name regionRange = workbook.Worksheets.Names[regionIndex];
                regionRange.RefersTo = "=Data!$A$5:$A$9";
                regionRange.Comment = "Valid sales regions for dropdown validation (non‑empty)";

                // Define named range "ProfitMargin"
                int profitIndex = workbook.Worksheets.Names.Add("ProfitMargin");
                Name profitRange = workbook.Worksheets.Names[profitIndex];
                profitRange.RefersTo = "=Data!$C$2:$C$13";
                profitRange.Comment = "Calculated profit margins (0%‑100%)";

                // Save the workbook
                string filePath = "NamedRangesWithComments.xlsx";
                workbook.Save(filePath);

                // Load the saved workbook to verify comments
                if (File.Exists(filePath))
                {
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Console.WriteLine("Named Range Comments after reload:");
                    foreach (Name name in loadedWorkbook.Worksheets.Names)
                    {
                        Console.WriteLine($"Name: {name.Text}, Comment: {name.Comment}");
                    }
                }
                else
                {
                    Console.WriteLine($"File not found: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeCommentDemo.Run();
        }
    }
}