using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create a custom number format style
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "#,##0.00"; // example custom format

        // Apply the style to a specific range (e.g., B2:B5)
        Worksheet sheet = workbook.Worksheets[0];
        Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:B5");
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // apply only the number format part
        range.ApplyStyle(customStyle, flag);

        // Configure text save options for CSV with tab delimiter
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
        txtOptions.Separator = '\t'; // set tab as the delimiter

        // Save the workbook as a tab‑delimited CSV file
        workbook.Save("output.tsv", txtOptions);
    }
}