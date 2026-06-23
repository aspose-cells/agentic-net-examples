using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsExamples
{
    public class EnablePrintGridlines
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data (optional, just to see the gridlines)
                worksheet.Cells["A1"].PutValue("Sample Data");
                worksheet.Cells["A2"].PutValue(123);

                // Enable printing of gridlines on the page
                worksheet.PageSetup.PrintGridlines = true;

                // Define output file path
                string outputPath = "PrintGridlinesEnabled.xlsx";

                // Save the workbook with the setting applied
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            EnablePrintGridlines.Run();
        }
    }
}