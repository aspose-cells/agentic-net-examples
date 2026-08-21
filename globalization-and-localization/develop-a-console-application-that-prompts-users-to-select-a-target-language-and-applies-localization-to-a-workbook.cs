// Title: C# Console App to Localize an Aspose.Cells Workbook via User‑Selected Language
// Description: A console program that prompts the user to choose a target language (USA, Germany, France, or Japan), maps the selection to a CountryCode enum, sets Workbook.Settings.LanguageCode and Workbook.Settings.Region, displays the derived CultureInfo, adds sample date and numeric cells that follow the locale's formatting rules, and saves the workbook as LocalizedWorkbook.xlsx.
// Keywords: Aspose.Cells | C# workbook localization | CountryCode enum | LanguageCode setting | Region setting | CultureInfo | console application | Excel regional formatting | date format localization | number format localization | user language selection
// Common Searches: Aspose.Cells set workbook language code C# | How to change workbook region with Aspose.Cells | Console program to apply locale to Excel file using Aspose | C# select country code for Excel localization | Aspose.Cells CultureInfo example
// Developer Intent: Build a C# console application that lets a user pick a language, applies the corresponding CountryCode to the workbook's LanguageCode and Region, demonstrates locale‑specific formatting, and saves the localized Excel file.
// Use Cases: Prompt the user for a language choice, convert it to a CountryCode, and configure workbook.Settings.LanguageCode and workbook.Settings.Region before saving. | Show the CultureInfo name derived from the selected CountryCode to verify the applied locale. | Insert sample cells with dates and numbers that automatically adopt the region‑specific formatting of the chosen language. | Extend the solution to support additional locales by mapping new CountryCode values.
// AI Prompts: Generate C# code that adds comprehensive validation and retry logic for user language selection in an Aspose.Cells localization console app. | Provide examples of customizing date and numeric format strings based on the selected CultureInfo within an Aspose.Cells workbook. | Explain how to programmatically enumerate all available CountryCode values in Aspose.Cells and present them as a dynamic menu for localization.

using System;
using System.Globalization;
using Aspose.Cells;

namespace WorkbookLocalizationDemo
{
    // A console program that prompts the user to choose a target language (USA, Germany, France, or Japan), maps the selection to a CountryCode enum, sets Workbook.Settings.LanguageCode and Workbook.Settings.Region, displays the derived CultureInfo, adds sample date and numeric cells that follow the locale's formatting rules, and saves the workbook as LocalizedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Prompt user to select a target language (CountryCode)
                Console.WriteLine("Select a target language for the workbook:");
                Console.WriteLine("1. United States (USA)");
                Console.WriteLine("2. Germany");
                Console.WriteLine("3. France");
                Console.WriteLine("4. Japan");
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
                    default:
                        Console.WriteLine("Invalid selection. Using default settings.");
                        break;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Apply localization settings
                workbook.Settings.LanguageCode = selectedCode; // UI language
                workbook.Settings.Region = selectedCode;       // Regional formatting

                // Show the CultureInfo derived from the selected region
                CultureInfo ci = workbook.Settings.CultureInfo;
                if (ci != null)
                {
                    Console.WriteLine($"CultureInfo for selected region: {ci.Name}");
                }

                // Add sample data to illustrate formatting
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Date value (formatted according to region)
                Cell dateCell = sheet.Cells["A1"];
                dateCell.PutValue(new DateTime(2023, 12, 31));
                Style dateStyle = dateCell.GetStyle();
                dateStyle.Custom = "yyyy-mm-dd";
                dateCell.SetStyle(dateStyle);

                // Number value (region‑specific separators)
                Cell numberCell = sheet.Cells["A2"];
                numberCell.PutValue(1234567.89);
                Style numberStyle = numberCell.GetStyle();
                numberStyle.Custom = "#,##0.00";
                numberCell.SetStyle(numberStyle);

                // Save the workbook
                string outputPath = "LocalizedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}' with language '{selectedCode}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
