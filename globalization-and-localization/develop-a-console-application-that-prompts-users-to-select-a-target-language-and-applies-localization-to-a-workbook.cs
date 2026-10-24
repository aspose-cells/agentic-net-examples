// Title: C# console program that reads a language code and localizes an Excel workbook with Aspose.Cells
// Common Searches: apply user-selected culture to Excel workbook using Aspose.Cells C# console | C# read language code from console and localize Excel file with Aspose.Cells | save localized Excel workbook with language code in filename using Aspose.Cells | handle invalid culture codes in Aspose.Cells C# application | set workbook Settings.CultureInfo based on user input Aspose.Cells
// Tags: Aspose.Cells workbook culture configuration | C# console user language input | localize Excel workbook Aspose.Cells | save workbook with language suffix | handle CultureNotFoundException Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// The console app prompts for a language code, loads 'input.xlsx' into an Aspose.Cells Workbook, assigns the appropriate CultureInfo to workbook.Settings.CultureInfo (fallback to invariant culture if the code is invalid), and saves the localized file as 'output_{languageCode}.xlsx'.
class Program
{
    static void Main()
    {
        // Prompt user for target language code (e.g., en, fr, de)
        Console.Write("Enter target language code (e.g., en, fr, de): ");
        string languageCode = Console.ReadLine()?.Trim();

        if (string.IsNullOrEmpty(languageCode))
        {
            Console.WriteLine("No language code entered. Application will exit.");
            return;
        }

        // Load the workbook from a predefined file
        // (Replace "input.xlsx" with the actual path if needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply localization by setting the workbook's culture info
        try
        {
            CultureInfo culture = new CultureInfo(languageCode);
            workbook.Settings.CultureInfo = culture;
        }
        catch (CultureNotFoundException)
        {
            Console.WriteLine($"Culture '{languageCode}' not found. Using invariant culture.");
            workbook.Settings.CultureInfo = CultureInfo.InvariantCulture;
        }

        // Save the localized workbook with a language-specific file name
        string outputFile = $"output_{languageCode}.xlsx";
        workbook.Save(outputFile);
        Console.WriteLine($"Workbook saved as '{outputFile}'.");
    }
}
