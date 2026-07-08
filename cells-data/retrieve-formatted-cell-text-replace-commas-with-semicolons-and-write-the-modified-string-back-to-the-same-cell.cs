using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Example data: a cell containing commas
        sheet.Cells["A1"].PutValue("Apple, Banana, Cherry");

        // Retrieve the formatted text from the cell
        string originalText = sheet.Cells["A1"].StringValue;

        // Replace all commas with semicolons
        string modifiedText = originalText.Replace(",", ";");

        // Write the modified string back to the same cell
        sheet.Cells["A1"].PutValue(modifiedText);

        // Save the workbook
        workbook.Save("Modified.xlsx");
    }
}