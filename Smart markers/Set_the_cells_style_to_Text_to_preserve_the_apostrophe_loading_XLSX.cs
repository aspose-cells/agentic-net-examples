using System;
using Aspose.Cells;

class PreserveApostropheExample
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Choose the cell that contains the value starting with an apostrophe
        // For demonstration, we use cell A1; replace with the actual address if different
        Cell cell = worksheet.Cells["A1"];

        // Create a new style and enable the QuotePrefix flag
        // This tells Excel that the cell's value starts with a single quote and should be treated as text
        Style textStyle = workbook.CreateStyle();
        textStyle.QuotePrefix = true;

        // Apply the style to the cell
        cell.SetStyle(textStyle);

        // Save the workbook; the apostrophe will be preserved as text in the saved file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}