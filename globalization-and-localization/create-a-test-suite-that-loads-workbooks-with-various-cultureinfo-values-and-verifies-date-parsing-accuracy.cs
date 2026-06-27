using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCultureInfoTests
{
    public class TestCultureInfoDateParsing
    {
        // Path for the temporary workbook used in all tests
        private const string TestFilePath = "DateParsingTest.xlsx";

        // The date value that will be written to the workbook
        private static readonly DateTime SampleDate = new DateTime(2023, 12, 31, 15, 45, 0);

        public static void Main()
        {
            // Step 1: Create a workbook with a single date cell and save it
            CreateSampleWorkbook();

            // Step 2: Define cultures to test
            string[] cultureNames = { "en-US", "de-DE", "fr-FR", "ja-JP", "ar-SA" };

            // Step 3: Load the workbook with each culture and verify date parsing
            foreach (string cultureName in cultureNames)
            {
                VerifyDateParsing(cultureName);
            }

            Console.WriteLine("All culture tests completed.");
        }

        private static void CreateSampleWorkbook()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Put the sample date into cell A1
            wb.Worksheets[0].Cells["A1"].PutValue(SampleDate);

            // Apply a date format so the cell is recognized as a date
            Style style = wb.CreateStyle();
            style.Custom = "yyyy-mm-dd hh:mm:ss";
            wb.Worksheets[0].Cells["A1"].SetStyle(style);

            // Save the workbook (XLSX format)
            wb.Save(TestFilePath, SaveFormat.Xlsx);
        }

        private static void VerifyDateParsing(string cultureName)
        {
            // Prepare LoadOptions with the specific CultureInfo
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo(cultureName);

            // Load the workbook using the options
            Workbook wb = new Workbook(TestFilePath, loadOptions);

            // Retrieve the date from cell A1
            Cell dateCell = wb.Worksheets[0].Cells["A1"];
            DateTime parsedDate = dateCell.DateTimeValue;

            // Verify that the parsed date matches the original sample date (ignoring Kind)
            bool datesMatch = parsedDate == SampleDate;

            // Also verify that the workbook's Settings.CultureInfo reflects the loaded culture
            CultureInfo workbookCulture = wb.Settings.CultureInfo;
            bool cultureMatches = workbookCulture != null && workbookCulture.Name.Equals(cultureName, StringComparison.OrdinalIgnoreCase);

            // Output the verification result
            Console.WriteLine($"Culture: {cultureName}");
            Console.WriteLine($"  Parsed Date: {parsedDate:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"  Expected Date: {SampleDate:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"  Dates Match: {datesMatch}");
            Console.WriteLine($"  Workbook Settings CultureInfo: {(workbookCulture != null ? workbookCulture.Name : "null")}");
            Console.WriteLine($"  CultureInfo Matches LoadOptions: {cultureMatches}");
            Console.WriteLine();
        }
    }
}