using System;
using Aspose.Cells;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (uses the provided create rule)
            Workbook workbook = new Workbook();

            // Access the workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Set the UI language of the workbook to German (Germany)
            settings.LanguageCode = CountryCode.Germany;

            // Optional: display the current language code
            Console.WriteLine("Workbook language set to: " + settings.LanguageCode);

            // Save the workbook to an XLSX file (uses the provided save rule)
            workbook.Save("Workbook_With_German_Language.xlsx");

            // Inform the user that the operation completed
            Console.WriteLine("Workbook saved successfully.");
        }
    }
}