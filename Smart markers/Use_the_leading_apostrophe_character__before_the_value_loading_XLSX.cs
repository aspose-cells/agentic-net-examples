using System;
using Aspose.Cells;

class LeadingApostropheDemo
{
    static void Main()
    {
        // Create a new workbook and enable QuotePrefixToStyle.
        // When this setting is true, strings that start with a single quote
        // will have the QuotePrefix style applied automatically.
        Workbook workbook = new Workbook();
        workbook.Settings.QuotePrefixToStyle = true;

        // Put a value that begins with a leading apostrophe.
        // The apostrophe is used in Excel to indicate that the following
        // text should be treated as a literal string.
        Cell cell = workbook.Worksheets[0].Cells["A1"];
        cell.PutValue("'SampleText");

        // Save the workbook as an XLSX file.
        string filePath = "LeadingApostrophe.xlsx";
        workbook.Save(filePath);

        // Load the saved XLSX file.
        Workbook loadedWorkbook = new Workbook(filePath);
        Cell loadedCell = loadedWorkbook.Worksheets[0].Cells["A1"];

        // Output the cell's string value and whether the QuotePrefix style is set.
        Console.WriteLine("Loaded cell value: " + loadedCell.StringValue);
        Console.WriteLine("QuotePrefix style applied: " + loadedCell.GetStyle().QuotePrefix);
    }
}