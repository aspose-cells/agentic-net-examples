// Title: C# – Set Workbook Default Font to Arial 10 and Apply to All Cells with Aspose.Cells
// Description: C# example that creates a Workbook, changes Workbook.DefaultStyle font to Arial 10, uses StyleFlag to apply the style to every cell in the first worksheet, adds sample data, and saves as XLSX.
// Keywords: Aspose.Cells | C# | default font | Arial 10 | Workbook.DefaultStyle | StyleFlag | ApplyStyle | apply font to all cells | Excel formatting | XLSX export
// Common Searches: Aspose.Cells set default font C# | apply default style to entire worksheet Aspose.Cells | change workbook font to Arial 10 .NET | C# code to set default workbook font Aspose | StyleFlag Font property Aspose.Cells example
// Developer Intent: Set the workbook’s default font to Arial 10 and ensure every cell inherits that formatting.
// Use Cases: Generate a new Excel report where all cells use the corporate Arial 10 font by default. | Retrofit an existing worksheet to enforce a uniform font without iterating each cell. | Create a template workbook that automatically applies the required font to any added data.
// AI Prompts: Provide C# code that sets Aspose.Cells Workbook.DefaultStyle font to Arial 10 and applies it to all cells. | Explain how StyleFlag.Font works with Worksheet.Cells.ApplyStyle to update the whole sheet. | Show how to change the default font to Times New Roman 12 for a specific range using Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace DefaultFontExample
{
    // C# example that creates a Workbook, changes Workbook.DefaultStyle font to Arial 10, uses StyleFlag to apply the style to every cell in the first worksheet, adds sample data, and saves as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Set the workbook's default font to Arial, size 10
            workbook.DefaultStyle.Font.Name = "Arial";
            workbook.DefaultStyle.Font.Size = 10;

            // Apply the default style to all existing cells in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            StyleFlag flag = new StyleFlag();
            flag.Font = true;               // Apply only font-related settings
            sheet.Cells.ApplyStyle(workbook.DefaultStyle, flag);

            // Add some sample data to verify the font is applied
            sheet.Cells["A1"].PutValue("Cell A1 uses Arial 10");
            sheet.Cells["B2"].PutValue("Cell B2 also uses Arial 10");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DefaultFontWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
