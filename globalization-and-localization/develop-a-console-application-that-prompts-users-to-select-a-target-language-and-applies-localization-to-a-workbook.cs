using System;
using Aspose.Cells;
using Aspose.Cells; // Ensure Aspose.Cells library is referenced

namespace WorkbookLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Display available languages (CountryCode enum values)
            Console.WriteLine("Select a target language for the workbook:");
            var values = Enum.GetValues(typeof(CountryCode));
            int index = 1;
            foreach (CountryCode code in values)
            {
                Console.WriteLine($"{index}. {code}");
                index++;
            }

            // Read user selection
            Console.Write("Enter the number corresponding to your choice: ");
            string input = Console.ReadLine();
            if (!int.TryParse(input, out int choice) || choice < 1 || choice > values.Length)
            {
                Console.WriteLine("Invalid selection. Exiting.");
                return;
            }

            // Map selection to CountryCode
            CountryCode selectedCode = (CountryCode)values.GetValue(choice - 1);
            Console.WriteLine($"You selected: {selectedCode}");

            // Create a new workbook and add sample data
            Workbook workbook = new Workbook(); // create workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Apply localization using WorkbookSettings.LanguageCode property
            workbook.Settings.LanguageCode = selectedCode;

            // Save the workbook
            string fileName = "LocalizedWorkbook.xlsx";
            workbook.Save(fileName); // save workbook

            Console.WriteLine($"Workbook saved as '{fileName}' with language code set to {selectedCode}.");
        }
    }
}