// Title: How to apply a bold font style to header cells in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a bold Font style and assigns it to the first‑row cells of an Aspose.Cells worksheet. | Write a C# snippet that populates header values and applies the bold style to each header cell before saving the workbook.
// Common Searches: Aspose.Cells C# set bold font for first row header cells | C# example applying style to specific cells in an Excel workbook with Aspose.Cells | How to create and reuse a bold style for header row using Aspose.Cells .NET
// Tags: apply bold style to header cells Aspose.Cells | create reusable cell style .NET Aspose.Cells | format first row as header Excel Aspose.Cells | save workbook with styled header Aspose.Cells C#

using Aspose.Cells;
using System;

// The program creates a new Workbook, defines a bold Font style, writes header values to the first row, applies the bold style to each header cell, and saves the file as HeaderBold.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Header values to populate
        string[] headers = { "ID", "Name", "Date" };

        // Create a style with bold font
        Style boldStyle = workbook.CreateStyle();
        boldStyle.Font.IsBold = true;

        // Apply the header values and bold style to the first row
        for (int col = 0; col < headers.Length; col++)
        {
            // Set header text
            Cell cell = sheet.Cells[0, col];
            cell.PutValue(headers[col]);

            // Apply bold style
            cell.SetStyle(boldStyle);
        }

        // Save the workbook to a file
        workbook.Save("HeaderBold.xlsx");
    }
}
