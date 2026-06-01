using System;
using System.IO;
using Aspose.Cells;

namespace DisableScientificNotationExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to an optional input workbook
                string inputPath = "input.xlsx";

                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Write a large number that would normally appear in scientific notation
                Cell cell = sheet.Cells["A1"];
                cell.PutValue(12345678901234567890.0);

                // Disable scientific notation by applying a custom number format
                Style style = cell.GetStyle();
                style.Custom = "0"; // plain integer format without scientific notation
                cell.SetStyle(style);

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}