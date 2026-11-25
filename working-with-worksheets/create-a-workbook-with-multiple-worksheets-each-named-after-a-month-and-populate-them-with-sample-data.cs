// Title: Create an Excel workbook with a separate worksheet for each month and add sample table data using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to add a worksheet for each month in a workbook | populate each monthly sheet with sample rows using Aspose.Cells .NET | set worksheet name to month name and auto‑fit columns with Aspose.Cells C# | save workbook with multiple month worksheets as .xlsx using Aspose.Cells | generate monthly Excel tabs programmatically in C# with Aspose.Cells
// Tags: add monthly worksheets Aspose.Cells C# | write sample table to each worksheet Aspose.Cells | auto‑fit columns programmatically Aspose.Cells | dynamic worksheet naming by month Aspose.Cells | save workbook as xlsx Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a new Workbook, adds twelve worksheets named after the months, writes a header row and five sample data rows to each sheet, auto‑fits the columns, and saves the file as MonthsData.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Array of month names
        string[] months = new string[]
        {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };

        // Loop through each month, add a worksheet and populate sample data
        foreach (string month in months)
        {
            // Add a new worksheet to the workbook
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet = workbook.Worksheets[sheetIndex];

            // Set the worksheet name to the month name
            sheet.Name = month;

            // Populate sample data (simple table with headers and a few rows)
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Description");
            sheet.Cells["C1"].PutValue("Value");

            // Sample rows
            for (int i = 1; i <= 5; i++)
            {
                int row = i + 1; // Data starts from row 2
                sheet.Cells[row, 0].PutValue(i); // ID
                sheet.Cells[row, 1].PutValue($"{month} Item {i}"); // Description
                sheet.Cells[row, 2].PutValue(i * 10); // Value
            }

            // Auto-fit columns for better appearance
            sheet.AutoFitColumns();
        }

        // Save the workbook to a file
        workbook.Save("MonthsData.xlsx");
    }
}
