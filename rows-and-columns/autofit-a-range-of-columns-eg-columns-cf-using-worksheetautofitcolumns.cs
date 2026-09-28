// Title: How to auto‑fit columns C to F in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, writes values to cells C1‑F1, and calls Worksheet.AutoFitColumns(2,5) to auto‑size those columns with Aspose.Cells. | Provide a step‑by‑step example of using Aspose.Cells to adjust column widths for a specific index range and then save the workbook.
// Common Searches: Aspose.Cells C# auto fit specific columns by index | Worksheet.AutoFitColumns example for columns C through F | How to programmatically adjust column width range in Excel using Aspose.Cells .NET | Auto‑size selected columns in an Excel file with Aspose.Cells C# code
// Tags: auto‑fit column range Aspose.Cells | adjust Excel column width by index | auto‑size selected columns C‑F | column width optimization .NET | Excel column auto‑sizing with Aspose

using System;
using Aspose.Cells;

namespace AutoFitColumnsExample
{
    // // Creates a new workbook, writes sample text into cells C1‑F1, auto‑fits columns C‑F using the AutoFitColumns method with index range 2‑5, and saves the file as AutoFitColumns_C_F.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data to columns C (index 2) through F (index 5)
            worksheet.Cells["C1"].PutValue("Short");
            worksheet.Cells["D1"].PutValue("Medium length text");
            worksheet.Cells["E1"].PutValue("Very very long text that needs column autofit");
            worksheet.Cells["F1"].PutValue("Another example with different length");

            // Auto‑fit columns C‑F (indices 2 to 5)
            worksheet.AutoFitColumns(2, 5);

            // Save the workbook
            workbook.Save("AutoFitColumns_C_F.xlsx");
        }
    }
}
