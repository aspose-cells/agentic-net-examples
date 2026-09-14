// Title: How to format Excel cells as percentages with one decimal place using Aspose.Cells for .NET (C#)
// AI Prompts: Create a custom style with the format string "0.0%" and apply it to a specific cell range using Aspose.Cells StyleFlag. | Generate an Excel workbook in C# where numeric values are automatically displayed as percentages with one decimal precision. | Save the workbook after applying the percentage style to cells A1 through A3 with Aspose.Cells.
// Common Searches: Aspose.Cells C# format cells as percentage with one decimal place | custom number format 0.0% using Aspose.Cells .NET | apply number format only to a range with StyleFlag in Aspose.Cells | C# create Excel file showing 12.3% instead of 0.123 using Aspose.Cells
// Tags: custom percentage number format Aspose.Cells | apply style to cell range Aspose.Cells C# | StyleFlag number format only Aspose.Cells | Excel percentage formatting .NET | Aspose.Cells generate workbook with custom format

using System;
using Aspose.Cells;

// The example creates a new workbook, inserts decimal values, defines a custom style with the format "0.0%" to display percentages with one decimal place, applies this style to the range A1:A3 using a StyleFlag that targets only the number format, and saves the file as PercentagesWithOneDecimal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Put some sample numeric values (e.g., 0.1234 = 12.34%).
        sheet.Cells["A1"].PutValue(0.1234);
        sheet.Cells["A2"].PutValue(0.5678);
        sheet.Cells["A3"].PutValue(0.9);

        // Define a custom number format that shows percentages with one decimal place.
        // "0.0%" means the value will be multiplied by 100 and displayed with one decimal.
        Style percentStyle = workbook.CreateStyle();
        percentStyle.Custom = "0.0%";

        // Apply the custom style to the range containing the numbers.
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true; // Apply only the number format part of the style.
        sheet.Cells.CreateRange("A1:A3").ApplyStyle(percentStyle, flag);

        // Save the workbook to a file.
        workbook.Save("PercentagesWithOneDecimal.xlsx");
    }
}
