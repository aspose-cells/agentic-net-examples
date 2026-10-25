// Title: Auto‑fit rows 15‑20 in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example to auto‑fit a specific range of rows | How to use Worksheet.AutoFitRows to adjust rows 15 through 20 | Auto‑fit multiple rows in an Excel file using Aspose.Cells .NET | Set row height automatically for rows 15‑20 with Aspose.Cells | Aspose.Cells C# code to auto‑fit rows based on content
// Tags: Worksheet.AutoFitRows C# sample | auto fit row range Aspose.Cells | adjust multiple row heights .NET | Excel row height automation Aspose.Cells | programmatic row auto‑fit using Aspose.Cells

using System;
using Aspose.Cells;

// This C# example creates a workbook, accesses the first worksheet, auto‑fits rows 15‑20 using Worksheet.AutoFitRows, and saves the result as AutoFitRows_Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Auto‑fit rows 15‑20 (zero‑based index: start at 14, total 6 rows)
        sheet.AutoFitRows(14, 6);

        // Save the workbook to a file
        workbook.Save("AutoFitRows_Output.xlsx");
    }
}
