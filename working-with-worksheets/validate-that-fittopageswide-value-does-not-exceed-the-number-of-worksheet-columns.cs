// Title: How to validate and adjust Worksheet.PageSetup.FitToPagesWide so it does not exceed Excel's 16,384 column limit with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if FitToPagesWide exceeds Excel's maximum column count | adjust worksheet PageSetup.FitToPagesWide when value is larger than 16384 columns in .NET | validate worksheet print scaling FitToPagesWide against Excel column limit using C# | set FitToPagesWide to the allowed maximum before saving a workbook with Aspose.Cells
// Tags: fit-to-pages-wide column limit validation | worksheet page setup scaling adjustment | aspocells column limit enforcement | c# worksheet print scaling check | excel max columns fit-to-pages-wide correction

using Aspose.Cells;
using System;

// The example loads a workbook, iterates through each worksheet, reads the PageSetup.FitToPagesWide property, compares it with Excel's maximum of 16,384 columns, resets the value if it exceeds the limit, and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the FitToPagesWide setting; 0 means automatic scaling
            int fitToPagesWide = sheet.PageSetup.FitToPagesWide;

            // Excel supports a maximum of 16,384 columns (XFD)
            const int maxColumns = 16384;

            // Validate that FitToPagesWide does not exceed the column limit
            if (fitToPagesWide > maxColumns)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}': FitToPagesWide ({fitToPagesWide}) exceeds the maximum column count ({maxColumns}). Adjusting to {maxColumns}.");
                sheet.PageSetup.FitToPagesWide = maxColumns;
            }
        }

        // Save the workbook after validation (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
