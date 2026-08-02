// Title: Apply a Custom Currency Number Format to a Named Range with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, defines a named range "CurrencyValues", builds a style with the custom format "$#,##0.00", applies only the number format to the range using StyleFlag, and saves the file as CurrencyValuesFormatted.xlsx.
// Keywords: Aspose.Cells | C# | custom number format | currency format | named range | StyleFlag | ApplyStyle | Excel workbook | .NET
// Common Searches: Aspose.Cells apply custom number format to named range | C# set currency format for a range in Excel | How to use StyleFlag to change only number format in Aspose.Cells | Apply style to named range Aspose.Cells .NET | Format cells with $#,##0.00 using Aspose.Cells
// Developer Intent: Apply a custom currency number format to every cell inside the named range "CurrencyValues".
// Use Cases: Standardize monetary values in financial reports by formatting a predefined range. | Generate invoices where total and tax columns share a consistent currency style. | Export data to Excel and ensure a specific column displays values with a dollar sign and two decimals.
// AI Prompts: Show how to change the custom format to "€#,##0.00" for the same named range. | Demonstrate applying font and border styles together with the currency format to a named range. | Explain how to modify the number format of an existing named range without recreating the style object.

using System;
using Aspose.Cells;

namespace ApplyCustomNumberFormatToNamedRange
{
    // Creates a workbook, defines a named range "CurrencyValues", builds a style with the custom format "$#,##0.00", applies only the number format to the range using StyleFlag, and saves the file as CurrencyValuesFormatted.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();                     // create
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample numeric data
                sheet.Cells["B2"].PutValue(1234.56);
                sheet.Cells["B3"].PutValue(7890.12);
                sheet.Cells["B4"].PutValue(345.67);

                // Define a named range "CurrencyValues" that covers the sample cells
                // RefersTo string must be in A1 style and include the sheet name
                int nameIndex = workbook.Worksheets.Names.Add("CurrencyValues");
                Name currencyName = workbook.Worksheets.Names[nameIndex];
                currencyName.RefersTo = "=Sheet1!$B$2:$B$4";

                // ---------- Create a style with a custom number format ----------
                Style style = workbook.CreateStyle();
                style.Custom = "$#,##0.00";                            // custom format

                // Specify that only the number format should be applied
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;                              // apply only number format

                // Retrieve the range by its name and apply the style
                Aspose.Cells.Range namedRange = workbook.Worksheets.GetRangeByName("CurrencyValues");
                namedRange.ApplyStyle(style, flag);                    // apply style to all cells in the named range

                // ---------- Save the workbook ----------
                workbook.Save("CurrencyValuesFormatted.xlsx");          // save
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
