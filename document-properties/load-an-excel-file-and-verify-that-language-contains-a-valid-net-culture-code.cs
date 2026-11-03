// Title: Check that a named range 'Language' in an Excel workbook holds a valid .NET culture identifier using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# read named range value and validate .NET culture code | how to confirm Excel 'Language' cell contains a valid culture identifier in C# | validate culture info from Excel named range using Aspose.Cells | C# load workbook and ensure language cell matches CultureInfo | detect invalid .NET culture codes in Excel with Aspose.Cells
// Tags: Aspose.Cells named range extraction C# | Excel culture identifier validation .NET | Workbook loading and property check Aspose.Cells | CultureInfo lookup for Excel cell value | language cell verification in workbook

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, retrieves the global named range "Language", reads its cell value, and uses System.Globalization.CultureInfo to determine whether the value is a valid .NET culture identifier, reporting the result.
class LanguageCultureValidator
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the named range "Language" from the workbook (global named ranges)
        Name languageName = workbook.Worksheets.Names["Language"];
        if (languageName == null)
        {
            Console.WriteLine("The named range 'Language' was not found.");
            return;
        }

        // Get the cell that the named range refers to
        Aspose.Cells.Range range;
        try
        {
            range = languageName.GetRange();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unable to resolve the range for 'Language': {ex.Message}");
            return;
        }

        if (range == null)
        {
            Console.WriteLine("Unable to resolve the range for 'Language'.");
            return;
        }

        Cell languageCell = sheet.Cells[range.FirstRow, range.FirstColumn];
        string languageValue = languageCell.StringValue?.Trim();

        if (string.IsNullOrEmpty(languageValue))
        {
            Console.WriteLine("The 'Language' cell is empty.");
            return;
        }

        // Verify that the value is a valid .NET culture code
        try
        {
            CultureInfo culture = CultureInfo.GetCultureInfo(languageValue);
            Console.WriteLine($"'{languageValue}' is a valid .NET culture code: {culture.DisplayName}");
        }
        catch (CultureNotFoundException)
        {
            Console.WriteLine($"'{languageValue}' is NOT a valid .NET culture code.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while validating culture: {ex.Message}");
        }
    }
}
