// Title: Convert culture‑specific numeric strings to numbers in Aspose.Cells for .NET
// Description: Demonstrates how to set workbook decimal and group separators (e.g., German format), insert strings like "1,23" and "4.567,89", run ConvertStringToNumericValue to turn convertible text into numeric values, verify the results, and save the workbook.
// Keywords: Aspose.Cells ConvertStringToNumericValue | culture specific number parsing .NET | German decimal separator | NumberDecimalSeparator Aspose.Cells | NumberGroupSeparator Aspose.Cells | C# spreadsheet numeric conversion | localized number format Excel | parse European numbers Aspose
// Common Searches: Aspose.Cells parse German numbers | Convert string cells to numeric with custom separators .NET | Set NumberDecimalSeparator in Aspose.Cells workbook | ConvertStringToNumericValue not handling group separator | C# example for culture specific numeric conversion in Excel
// Developer Intent: The developer needs to transform text cells that use locale‑specific decimal and thousands separators into true numeric values without affecting non‑numeric entries.
// Use Cases: Import CSV files containing European number formats and convert them for calculations. | Load a workbook created in a localized environment, apply custom separators, and ensure formulas evaluate correctly. | Validate user‑entered numeric strings in a spreadsheet, converting only valid entries while preserving invalid text.
// AI Prompts: Show C# code that converts culture‑specific numeric strings to numbers after setting custom decimal and group separators in Aspose.Cells. | Explain how to detect and convert only convertible string cells while leaving non‑numeric values unchanged. | Provide guidance on handling conversion failures or exceptions when using ConvertStringToNumericValue.

using System;
using Aspose.Cells;

// Demonstrates how to set workbook decimal and group separators (e.g., German format), insert strings like "1,23" and "4.567,89", run ConvertStringToNumericValue to turn convertible text into numeric values, verify the results, and save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set culture‑specific decimal and group separators (e.g., German format)
        wb.Settings.NumberDecimalSeparator = ',';
        wb.Settings.NumberGroupSeparator = '.';

        // Access the first worksheet's cells
        Cells cells = wb.Worksheets[0].Cells;

        // Insert numeric strings that use the specified separators
        cells["A1"].PutValue("1,23");          // Simple decimal
        cells["A2"].PutValue("4.567,89");      // Group separator with decimal
        cells["A3"].PutValue("not a number"); // Non‑numeric string

        // Convert all convertible string values to numeric values
        cells.ConvertStringToNumericValue();

        // Output the converted values to verify correct parsing
        Console.WriteLine("A1 numeric value: " + cells["A1"].DoubleValue);
        Console.WriteLine("A2 numeric value: " + cells["A2"].DoubleValue);
        Console.WriteLine("A3 type after conversion: " + cells["A3"].Value.GetType());

        // Save the workbook
        wb.Save("ConvertedNumbers.xlsx");
    }
}
