using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadOptionsCultureInfoErrorHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Path to a sample workbook (created if it does not exist)
                string sourcePath = "sample.xlsx";
                EnsureSampleWorkbookExists(sourcePath);

                // Locale identifier that might be unsupported
                string localeId = "xx-XX"; // intentionally invalid

                // Prepare LoadOptions with error handling for CultureInfo
                LoadOptions loadOptions = new LoadOptions();

                try
                {
                    // Attempt to assign the requested CultureInfo
                    loadOptions.CultureInfo = new CultureInfo(localeId);
                }
                catch (CultureNotFoundException ex)
                {
                    // Unsupported locale – log and fall back to InvariantCulture
                    Console.WriteLine($"Locale '{localeId}' is not supported: {ex.Message}");
                    loadOptions.CultureInfo = CultureInfo.InvariantCulture;
                    Console.WriteLine("Falling back to InvariantCulture.");
                }

                // Verify source file exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the workbook using the prepared LoadOptions
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Demonstrate that the culture is applied (e.g., number formatting)
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A2"];
                cell.PutValue(1234.56); // value to be formatted

                // Apply a numeric style to show culture‑specific decimal separator
                Style style = workbook.CreateStyle();
                style.Number = 2; // two decimal places
                cell.SetStyle(style);

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper to create a minimal workbook if the sample file is missing
        private static void EnsureSampleWorkbookExists(string path)
        {
            if (!File.Exists(path))
            {
                Workbook wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue("Sample");
                wb.Save(path);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadOptionsCultureInfoErrorHandlingDemo.Run();
        }
    }
}