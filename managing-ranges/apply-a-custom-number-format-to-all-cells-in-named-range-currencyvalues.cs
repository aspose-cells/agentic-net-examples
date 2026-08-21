// Title: Apply a custom Euro number format to a named range with Aspose.Cells for .NET
// Description: Creates a workbook, defines the named range "CurrencyValues" (A1:A3), builds a Style with the Euro custom format "_-€ #,##0.00;[Red]_-€ -#,##0.00", uses a StyleFlag to affect only the number format, retrieves the range reference, and applies the style to every cell in the named range before saving the file.
// Keywords: Aspose.Cells custom number format | C# named range formatting | apply style to range Aspose.Cells | Euro currency format Excel | StyleFlag number format only | Aspose.Cells .NET example
// Common Searches: Aspose.Cells apply custom number format to named range | C# set Euro currency format for a range in Excel | How to use StyleFlag with Aspose.Cells | Create and format named range Aspose.Cells .NET | Apply custom number format to multiple cells Aspose
// Developer Intent: Apply a Euro‑style custom number format to every cell inside the "CurrencyValues" named range using Aspose.Cells for .NET.
// Use Cases: Standardize financial column appearance in generated reports by applying a single Euro format to a predefined named range. | Reuse the "CurrencyValues" named range across worksheets to ensure consistent currency display with red negatives. | Build a template workbook where any cell added to the "CurrencyValues" range automatically inherits the custom Euro formatting.
// AI Prompts: Generate C# code that creates a named range "CurrencyValues" and applies the custom number format "_-€ #,##0.00;[Red]_-€ -#,##0.00" using Aspose.Cells. | Explain how to modify the custom format string for different currencies while applying it to a named range with Aspose.Cells. | Provide a step‑by‑step guide to retrieve a named range reference and use StyleFlag to change only the number format in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace ApplyCustomNumberFormat
{
    // Creates a workbook, defines the named range "CurrencyValues" (A1:A3), builds a Style with the Euro custom format "_-€ #,##0.00;[Red]_-€ -#,##0.00", uses a StyleFlag to affect only the number format, retrieves the range reference, and applies the style to every cell in the named range before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample numeric data that will belong to the named range
                sheet.Cells["A1"].PutValue(1234.56);
                sheet.Cells["A2"].PutValue(7890.12);
                sheet.Cells["A3"].PutValue(345.67);

                // Define a named range called "CurrencyValues" covering A1:A3
                int nameIdx = workbook.Worksheets.Names.Add("CurrencyValues");
                Name currencyRange = workbook.Worksheets.Names[nameIdx];
                // RefersTo must include sheet name and absolute references
                currencyRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Create a style with the desired custom number format
                Style customStyle = workbook.CreateStyle();
                customStyle.Custom = "_-€ #,##0.00;[Red]_-€ -#,##0.00";

                // Configure a StyleFlag to apply only the number format part
                StyleFlag flag = new StyleFlag();
                flag.NumberFormat = true;

                // Retrieve the range reference string from the named range (A1:A3)
                string rangeRef = currencyRange.GetRefersTo(false, false); // e.g., "=Sheet1!$A$1:$A$3"
                if (rangeRef.StartsWith("="))
                    rangeRef = rangeRef.Substring(1); // remove leading '='

                // Create a Range object based on the reference (sheet name is optional here because we are on the same sheet)
                Aspose.Cells.Range targetRange = sheet.Cells.CreateRange(rangeRef);

                // Apply the custom number format style to the entire named range
                targetRange.ApplyStyle(customStyle, flag);

                // Save the workbook to verify the formatting
                workbook.Save("CurrencyValuesFormatted.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
