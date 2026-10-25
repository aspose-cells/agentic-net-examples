// Title: Auto‑fit a single Excel column based on its content with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# auto fit column width for column A only | How to resize a specific Excel column using Worksheet.AutoFitColumn in .NET | C# code to auto‑adjust column width after writing data with Aspose.Cells | Example of using Worksheet.AutoFitColumn to fit column content in a workbook | Programmatically auto‑size a single column in an Excel file using Aspose.Cells API
// Tags: Worksheet.AutoFitColumn usage .NET | single column sizing Aspose.Cells | Excel column auto sizing C# API | dynamic column sizing based on cell content | programmatic column resizing with Aspose.Cells

using System;
using Aspose.Cells;

namespace AutoFitColumnExample
{
    // The sample creates a new Workbook, writes three strings of varying length into cells A1‑A3 of the first worksheet, calls Worksheet.AutoFitColumn(0) to automatically adjust column A's width to fit its content, and saves the file as AutoFitColumnResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (uses the create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells in column A (index 0) with sample data
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["A2"].PutValue("A bit longer text");
            sheet.Cells["A3"].PutValue("This is a considerably longer piece of text to demonstrate auto‑fit.");

            // Auto‑fit column A based on its content (column index 0)
            sheet.AutoFitColumn(0);

            // Save the workbook (uses the save rule)
            workbook.Save("AutoFitColumnResult.xlsx");
        }
    }
}
