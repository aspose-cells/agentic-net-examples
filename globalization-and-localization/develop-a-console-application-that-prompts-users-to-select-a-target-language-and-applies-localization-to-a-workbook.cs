using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Prompt user to select a target language (country code)
            Console.WriteLine("Select a target language for the workbook:");
            Console.WriteLine("1 - United States (USA)");
            Console.WriteLine("2 - Germany");
            Console.WriteLine("3 - France");
            Console.WriteLine("4 - Japan");
            Console.WriteLine("5 - China");
            Console.Write("Enter the number of your choice: ");

            string input = Console.ReadLine();
            CountryCode selectedCode = CountryCode.Default;

            switch (input)
            {
                case "1":
                    selectedCode = CountryCode.USA;
                    break;
                case "2":
                    selectedCode = CountryCode.Germany;
                    break;
                case "3":
                    selectedCode = CountryCode.France;
                    break;
                case "4":
                    selectedCode = CountryCode.Japan;
                    break;
                case "5":
                    selectedCode = CountryCode.China;
                    break;
                default:
                    Console.WriteLine("Invalid selection. Using default language settings.");
                    break;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Apply the selected language code to workbook settings
            if (selectedCode != CountryCode.Default)
            {
                workbook.Settings.LanguageCode = selectedCode;
                // Optionally set the region to match the language
                workbook.Settings.Region = selectedCode;
            }

            // Add sample data to demonstrate the workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Localized Workbook");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["A3"].PutValue(12345.67);

            // Save the workbook
            string outputPath = "LocalizedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}' with language code: {workbook.Settings.LanguageCode}");
        }
    }
}