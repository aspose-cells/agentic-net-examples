using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data containing commas
        worksheet.Cells["A1"].PutValue("Apple, Banana, Cherry");

        // Retrieve the formatted text from the cell
        Cell cell = worksheet.Cells["A1"];
        string formattedText = cell.StringValue; // Gets the displayed string value

        // Replace commas with semicolons
        string modifiedText = formattedText.Replace(",", ";");

        // Write the modified string back to the same cell
        cell.PutValue(modifiedText);

        // Save the workbook
        workbook.Save("ModifiedCell.xlsx");
    }
}