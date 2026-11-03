// Title: Import an ArrayList of decimal numbers into an Excel worksheet and round each cell to two decimal places with Aspose.Cells for .NET
// Common Searches: asp.net import ArrayList of decimals into Excel using Aspose.Cells | round numbers to two decimal places in an Aspose.Cells workbook C# | apply Math.Round to cell values after importing data with Aspose.Cells | save Excel file after rounding decimal values with Aspose.Cells .NET | convert ArrayList to worksheet cells Aspose.Cells C# example
// Tags: import decimal ArrayList into worksheet Aspose.Cells | round numeric cells to two decimal places C# | apply Math.Round MidpointRounding in Aspose.Cells | save workbook as .xlsx using Aspose.Cells | populate Excel column from ArrayList C#

using System;
using System.Collections;
using Aspose.Cells;

// The example creates a Workbook, imports a list of decimal numbers from an ArrayList into column A of the first worksheet, rounds each cell value to two decimal places using Math.Round with MidpointRounding.AwayFromZero, and saves the result as RoundedValues.xlsx.
class Program
{
    static void Main()
    {
        // Prepare the source data: an ArrayList of decimal numbers
        ArrayList decimalValues = new ArrayList
        {
            12.3456m,
            78.9012m,
            3.14159m,
            0.9999m,
            123.4567m
        };

        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Import data into the worksheet starting at cell A1
        for (int i = 0; i < decimalValues.Count; i++)
        {
            // Write the original decimal value to the cell
            cells[i, 0].Value = decimalValues[i];
        }

        // Round each cell value to two decimal places
        for (int i = 0; i < decimalValues.Count; i++)
        {
            // Retrieve the current value as decimal
            decimal original = Convert.ToDecimal(cells[i, 0].Value);
            // Round to two decimal places
            decimal rounded = Math.Round(original, 2, MidpointRounding.AwayFromZero);
            // Write the rounded value back to the cell
            cells[i, 0].Value = rounded;
        }

        // Save the workbook to a file
        workbook.Save("RoundedValues.xlsx");
    }
}
