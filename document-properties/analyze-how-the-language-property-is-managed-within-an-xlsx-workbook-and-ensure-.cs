using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsLanguageDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------- Create a new workbook -------------------
            Workbook workbook = new Workbook();

            // Set the UI language of the workbook (affects how Excel displays UI elements)
            workbook.Settings.LanguageCode = CountryCode.Germany; // German (Germany)

            // Set the document's language property (metadata)
            workbook.BuiltInDocumentProperties.Language = "de-DE";

            // Set regional settings (affects number/date formatting)
            workbook.Settings.Region = CountryCode.Germany;

            // After setting Region, CultureInfo reflects the corresponding culture
            CultureInfo ci = workbook.Settings.CultureInfo;
            Console.WriteLine($"CultureInfo after setting Region to Germany: {ci.Name}");

            // Demonstrate custom globalization settings (e.g., boolean strings)
            SettableGlobalizationSettings customGlobalization = new SettableGlobalizationSettings();
            customGlobalization.SetBooleanValueString(true, "WAHR");
            customGlobalization.SetBooleanValueString(false, "FALSCH");
            workbook.Settings.GlobalizationSettings = customGlobalization;

            // Add sample data to show formatting and boolean strings
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue(true);   // Will display "WAHR"
            ws.Cells["A2"].PutValue(false);  // Will display "FALSCH"
            ws.Cells["B1"].PutValue(1234567.89); // Number formatting will use German separators

            // Save the workbook (XLSX)
            string filePath = "LocalizedWorkbook.xlsx";
            workbook.Save(filePath);
            Console.WriteLine($"Workbook saved to {filePath}");

            // ------------------- Load the workbook with different language settings -------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguageCode = CountryCode.France; // UI language for loading
            loadOptions.CultureInfo = new CultureInfo("fr-FR"); // French culture for parsing numbers/dates

            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Verify loaded settings
            Console.WriteLine($"Loaded Workbook Settings.LanguageCode: {loadedWorkbook.Settings.LanguageCode}");
            Console.WriteLine($"Loaded Workbook BuiltInDocumentProperties.Language: {loadedWorkbook.BuiltInDocumentProperties.Language}");
            Console.WriteLine($"Loaded Workbook Settings.Region: {loadedWorkbook.Settings.Region}");
            Console.WriteLine($"Loaded Workbook Settings.CultureInfo: {loadedWorkbook.Settings.CultureInfo.Name}");

            // Show that boolean values respect the custom globalization set earlier
            Worksheet loadedWs = loadedWorkbook.Worksheets[0];
            Console.WriteLine($"Cell A1 (bool true) displayed as: {loadedWs.Cells["A1"].StringValue}");
            Console.WriteLine($"Cell A2 (bool false) displayed as: {loadedWs.Cells["A2"].StringValue}");
            Console.WriteLine($"Cell B1 (number) displayed as: {loadedWs.Cells["B1"].StringValue}");

            // Save the loaded workbook to a new file to confirm no loss of settings
            string loadedFilePath = "LocalizedWorkbook_Loaded.xlsx";
            loadedWorkbook.Save(loadedFilePath);
            Console.WriteLine($"Loaded workbook saved to {loadedFilePath}");
        }
    }
}