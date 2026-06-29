using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the merged workbook (use the provided load rule)
        Workbook workbook = new Workbook("MergedWorkbook.xlsx"); // <-- load rule placeholder

        // Apply AutoFit to all columns in each worksheet, handling merged cells appropriately
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Configure AutoFitterOptions to expand merged cells line‑by‑line and respect wrapped text
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // AutoFit columns using the options
            sheet.AutoFitColumns(options);
        }

        // Save the workbook after autofitting (use the provided save rule)
        workbook.Save("MergedWorkbook_AutoFitted.xlsx"); // <-- save rule placeholder
    }
}
// Author: Aspose.Cells .NET example – applies column AutoFit to a merged workbook.