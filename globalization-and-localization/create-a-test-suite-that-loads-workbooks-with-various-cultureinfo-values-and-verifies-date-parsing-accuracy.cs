using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureInfoTest
{
    class Program
    {
        static void Main()
        {
            // Path for the temporary workbook
            const string workbookPath = "DateSample.xlsx";

            // 1. Create a workbook with a known date value and save it
            CreateSampleWorkbook(workbookPath);

            // 2. Define cultures to test
            var cultures = new Dictionary<string, CultureInfo>
            {
                { "en-US", new CultureInfo("en-US") }, // MM/dd/yyyy
                { "de-DE", new CultureInfo("de-DE") }, // dd.MM.yyyy
                { "fr-FR", new CultureInfo("fr-FR") }  // dd/MM/yyyy
            };

            // Expected date (January 15, 2023)
            DateTime expectedDate = new DateTime(2023, 1, 15);

            // 3. Load the workbook with each culture and verify date parsing
            foreach (var kvp in cultures)
            {
                string cultureName = kvp.Key;
                CultureInfo culture = kvp.Value;

                // Use LoadOptions with the specific CultureInfo (rule: LoadOptions.CultureInfo)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                loadOptions.CultureInfo = culture;

                // Load the workbook (rule: load)
                Workbook wb = new Workbook(workbookPath, loadOptions);
                Worksheet sheet = wb.Worksheets[0];
                Cell dateCell = sheet.Cells["A1"];

                // Retrieve the parsed date
                DateTime parsedDate = dateCell.DateTimeValue;

                // Verify that the parsed date matches the expected date
                if (parsedDate != expectedDate)
                {
                    Console.WriteLine($"[FAIL] Culture {cultureName}: Parsed date {parsedDate:d} does not match expected {expectedDate:d}");
                }
                else
                {
                    Console.WriteLine($"[PASS] Culture {cultureName}: Parsed date correctly as {parsedDate:d}");
                }
            }

            // Clean up (optional)
            // System.IO.File.Delete(workbookPath);
        }

        // Helper method to create a workbook containing a date in cell A1 and save it
        static void CreateSampleWorkbook(string path)
        {
            // Create a new workbook (rule: create)
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Put a known date value into cell A1
            DateTime sampleDate = new DateTime(2023, 1, 15);
            sheet.Cells["A1"].PutValue(sampleDate);

            // Apply a date format (optional, does not affect parsing)
            Style style = sheet.Cells["A1"].GetStyle();
            style.Custom = "mm/dd/yyyy";
            sheet.Cells["A1"].SetStyle(style);

            // Save the workbook (rule: save)
            wb.Save(path, SaveFormat.Xlsx);
        }
    }
}