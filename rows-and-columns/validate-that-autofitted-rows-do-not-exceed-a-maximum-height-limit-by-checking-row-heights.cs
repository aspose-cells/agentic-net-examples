// Title: C# – Validate AutoFitRows Does Not Exceed Max Row Height with Aspose.Cells
// Description: Creates a workbook, adds wrapped text, sets an initial row height, defines a maximum height, and uses AutoFitterOptions.MaxRowHeight to auto‑fit rows. The code then checks each populated row, reports whether its height stays within the limit, and saves the file.
// Keywords: Aspose.Cells | AutoFitRows | MaxRowHeight | C# | row height validation | limit row height | wrap text Excel | Excel automation .NET | row height constraint | AutoFitterOptions example
// Common Searches: Aspose.Cells limit row height when auto fitting | C# check row height after AutoFitRows | MaxRowHeight option usage Aspose.Cells | validate Excel row height does not exceed maximum | auto‑fit rows with height cap .NET
// Developer Intent: Confirm that rows auto‑fitted by Aspose.Cells stay below a predefined height threshold.
// Use Cases: Prevent excessively tall rows in reports that contain wrapped text. | Enforce a uniform row‑height ceiling before printing or exporting to PDF. | Audit generated worksheets and flag rows that violate height constraints.
// AI Prompts: Show how to throw a custom exception when a row exceeds MaxRowHeight instead of writing to console. | Demonstrate setting both MaxRowHeight and MaxColumnWidth in AutoFitterOptions for a worksheet. | Explain how to retrieve the actual row height in points after AutoFitRows and compare it to a configurable limit.

using System;
using Aspose.Cells;

// Creates a workbook, adds wrapped text, sets an initial row height, defines a maximum height, and uses AutoFitterOptions.MaxRowHeight to auto‑fit rows. The code then checks each populated row, reports whether its height stays within the limit, and saves the file.
class ValidateAutoFitRowHeight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add long wrapped text to cause row height expansion
        sheet.Cells["A1"].PutValue("This is a very long text that should cause the row to expand beyond the limit when auto‑fitted.");
        Style style = sheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        sheet.Cells["A1"].SetStyle(style);

        // Set an initial small row height
        sheet.Cells.SetRowHeight(0, 10);

        // Define the maximum allowed row height (in points)
        double maxHeight = 30.0;

        // Create AutoFitterOptions with MaxRowHeight limit
        AutoFitterOptions options = new AutoFitterOptions
        {
            MaxRowHeight = maxHeight,
            OnlyAuto = true
        };

        // Auto‑fit rows using the options (rows will not exceed maxHeight)
        sheet.AutoFitRows(options);

        // Validate that each row height does not exceed the maximum limit
        int lastDataRow = sheet.Cells.MaxDataRow; // index of the last row containing data
        for (int rowIndex = 0; rowIndex <= lastDataRow; rowIndex++)
        {
            double actualHeight = sheet.Cells.GetRowHeight(rowIndex);
            if (actualHeight > maxHeight)
            {
                Console.WriteLine($"Row {rowIndex} height {actualHeight} exceeds the maximum allowed {maxHeight}.");
            }
            else
            {
                Console.WriteLine($"Row {rowIndex} height {actualHeight} is within the limit.");
            }
        }

        // Save the workbook
        workbook.Save("ValidatedAutoFitRows.xlsx");
    }
}
