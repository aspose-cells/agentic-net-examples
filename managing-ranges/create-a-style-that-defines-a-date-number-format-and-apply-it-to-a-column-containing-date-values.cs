// Title: Create a date number format style and apply it to an entire column with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a Style with a short date number format (built‑in ID 14 or a custom pattern) and applies it to column A using StyleFlag so only the number format changes. | Demonstrate how to apply the same date style to every cell in a column without affecting other cell attributes in an Aspose.Cells workbook.
// Common Searches: aspocells c# apply built‑in short date format to whole column | how to use StyleFlag to change only number format in Aspose.Cells | set custom date format mm-dd-yyyy for column A using Aspose.Cells .NET | apply same style to entire column in Excel with Aspose.Cells C# example | Aspose.Cells format column as date without altering other cell properties
// Tags: Aspose.Cells date style column | C# StyleFlag number format only | custom short date format Aspose.Cells | apply style to whole column Excel .NET | built‑in date format ID 14 Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a workbook, inserts three DateTime values into column A, defines a style with a short date number format (built‑in ID 14 or a custom pattern), uses StyleFlag to apply only the number format to the entire first column, and saves the file as DateStyleExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample date values in column A (index 0)
        sheet.Cells["A1"].PutValue(DateTime.Now);
        sheet.Cells["A2"].PutValue(DateTime.Now.AddDays(1));
        sheet.Cells["A3"].PutValue(DateTime.Now.AddDays(2));

        // Create a style for date formatting
        Style dateStyle = workbook.CreateStyle();

        // Use built‑in short date format (mm-dd-yyyy)
        dateStyle.Number = 14; // Built‑in date format ID

        // If a custom format is preferred, uncomment the next line:
        // dateStyle.Custom = "mm-dd-yyyy";

        // Specify that only the number format should be applied
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the entire first column (A)
        sheet.Cells.Columns[0].ApplyStyle(dateStyle, flag);

        // Save the workbook
        workbook.Save("DateStyleExample.xlsx");
    }
}
