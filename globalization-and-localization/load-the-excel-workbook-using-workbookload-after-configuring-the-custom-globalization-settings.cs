// Title: Load an Excel workbook with German locale using Aspose.Cells LoadOptions (C#)
// Description: Shows how to set LoadOptions.CultureInfo to "de-DE" so Aspose.Cells interprets commas as decimal separators, loads the workbook, reads a cell value, and optionally saves the file.
// Keywords: Aspose.Cells LoadOptions | CultureInfo de-DE | German locale Excel .NET | custom globalization Aspose.Cells | comma decimal separator | load workbook with culture | C# Aspose.Cells localization
// Common Searches: Aspose.Cells load workbook with German culture | Set CultureInfo in LoadOptions C# | Excel decimal separator comma Aspose.Cells | Configure localization for Excel import .NET | LoadOptions CultureInfo example
// Developer Intent: Load a workbook while applying a specific CultureInfo via LoadOptions.
// Use Cases: Import German‑formatted spreadsheets where numbers use commas | Parse dates and numbers according to locale settings during data migration | Generate reports for German users with correct number formatting | Batch‑process files that contain mixed regional formats
// AI Prompts: Show code to load an Excel file with French locale using Aspose.Cells LoadOptions in C#. | How to set CultureInfo to Japanese and read date cells with Aspose.Cells. | Explain global configuration of Aspose.Cells culture for all loads in a .NET application. | Provide an example of using LoadOptions to handle Arabic numerals in Excel.

using System;
using System.Globalization;
using Aspose.Cells;

// Shows how to set LoadOptions.CultureInfo to "de-DE" so Aspose.Cells interprets commas as decimal separators, loads the workbook, reads a cell value, and optionally saves the file.
class Program
{
    static void Main()
    {
        // Configure load options with a custom culture (German in this example)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("de-DE"); // German uses comma as decimal separator

        // Load the workbook using the constructor that accepts LoadOptions
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile, loadOptions);

        // Demonstrate that the culture is applied by reading a cell value
        string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Cell A1 value with German culture: " + cellValue);

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}
