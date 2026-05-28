using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyCellLengthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Ensure Excel restriction checking is enabled (default true)
                workbook.Settings.CheckExcelRestriction = true;

                // Access the first worksheet and its cells
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Prepare a test string longer than 255 characters
                string longText = new string('A', 300);

                // Put the long string into cell A1
                cells["A1"].PutValue(longText);

                // Verify the length of the cell's string value
                int length = cells["A1"].StringValue.Length;

                // If compatibility mode (Excel 2003) is considered, enforce 255‑character limit
                if (length > 255)
                {
                    Console.WriteLine($"Warning: Cell A1 contains {length} characters, which exceeds the 255‑character limit for Excel 2003 compatibility.");

                    // Truncate the value to meet the limit
                    string truncated = cells["A1"].StringValue.Substring(0, 255);
                    cells["A1"].PutValue(truncated);
                    Console.WriteLine("Cell A1 value has been truncated to 255 characters.");
                }
                else
                {
                    Console.WriteLine($"Cell A1 length is within the limit: {length} characters.");
                }

                // Save the workbook (lifecycle: save)
                string outputPath = "VerifiedCellLength.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            VerifyCellLengthDemo.Run();
        }
    }
}