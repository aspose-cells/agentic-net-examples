// Title: How to validate and correct FitToPagesWide so it doesn't exceed the column count in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if FitToPagesWide is larger than the number of columns in a worksheet | programmatically limit FitToPagesWide to the maximum column count with Aspose.Cells .NET | adjust worksheet PageSetup.FitToPagesWide based on actual column count using Aspose.Cells | prevent FitToPagesWide overflow when printing Excel files with Aspose.Cells in C# | iterate worksheets and reset FitToPagesWide to column count Aspose.Cells example
// Tags: page setup FitToPagesWide limit Aspose.Cells | worksheet column count retrieval C# Aspose.Cells | adjust print scaling to column count Aspose.Cells | validate FitToPagesWide against MaxColumn Aspose.Cells | automated worksheet print settings correction Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates each worksheet, calculates the used column count via Cells.MaxColumn, compares it with the worksheet's PageSetup.FitToPagesWide, logs a warning if the setting exceeds the column count, and resets FitToPagesWide to the column count before saving the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the number of columns that contain data.
            // MaxColumn is zero‑based, so add 1 to get the count.
            int columnCount = sheet.Cells.MaxColumn + 1;

            // Retrieve the current FitToPagesWide setting.
            int fitToPagesWide = sheet.PageSetup.FitToPagesWide;

            // Validate that FitToPagesWide does not exceed the column count.
            if (fitToPagesWide > columnCount)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\": FitToPagesWide ({fitToPagesWide}) exceeds column count ({columnCount}). Adjusting to {columnCount}.");
                // Adjust the setting to the maximum allowed value.
                sheet.PageSetup.FitToPagesWide = columnCount;
            }
        }

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
