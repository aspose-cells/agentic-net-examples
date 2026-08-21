// Title: C# – Aspose.Cells: Create a workbook with a worksheet for each month and add sample data
// Description: This Aspose.Cells for .NET example builds a new Workbook, renames the first sheet to "January", adds eleven more worksheets named February‑December, and fills every sheet with a header (Date, Description, Amount) plus five sample rows for the first days of the month. The file is saved as MonthsWorkbook.xlsx.
// Keywords: Aspose.Cells | C# | .NET | create workbook | add worksheets | month worksheets | rename default sheet | populate cells | sample data | save as xlsx | Excel automation | multiple sheets | monthly template
// Common Searches: Aspose.Cells add worksheet for each month C# | populate multiple sheets with sample data Aspose.Cells | rename default worksheet Aspose.Cells .NET | save workbook with 12 month sheets Aspose.Cells | example code Aspose.Cells monthly worksheets
// Developer Intent: Generate an Excel file containing twelve month‑named worksheets, each pre‑filled with a header and sample rows.
// Use Cases: Build a monthly expense tracker where each month has its own sheet. | Create a test workbook with placeholder data for performance or UI testing. | Generate a calendar‑based data entry template that separates entries by month.
// AI Prompts: Write C# code using Aspose.Cells to add worksheets named January‑December and insert a header and sample rows into each. | Show how to apply bold formatting to the header row and auto‑fit columns for the month worksheets created with Aspose.Cells. | Explain how to replace the static sample rows with data from a collection or database for each month sheet.

using System;
using Aspose.Cells;

// This Aspose.Cells for .NET example builds a new Workbook, renames the first sheet to "January", adds eleven more worksheets named February‑December, and fills every sheet with a header (Date, Description, Amount) plus five sample rows for the first days of the month. The file is saved as MonthsWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Month names to be used as worksheet names
        string[] months = new string[]
        {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };

        // Rename the default worksheet to the first month and populate it
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.Name = months[0];
        PopulateWorksheet(firstSheet, months[0]);

        // Add remaining months as new worksheets and populate them
        for (int i = 1; i < months.Length; i++)
        {
            Worksheet ws = workbook.Worksheets.Add(months[i]);
            PopulateWorksheet(ws, months[i]);
        }

        // Save the workbook to a file
        workbook.Save("MonthsWorkbook.xlsx", SaveFormat.Xlsx);
    }

    // Adds sample header and data rows to a worksheet
    static void PopulateWorksheet(Worksheet sheet, string monthName)
    {
        // Header row
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Description");
        sheet.Cells["C1"].PutValue("Amount");

        // Sample data rows (first five days of the month)
        for (int day = 1; day <= 5; day++)
        {
            int row = day + 1; // Data starts from row 2
            sheet.Cells[$"A{row}"].PutValue($"{monthName} {day}");
            sheet.Cells[$"B{row}"].PutValue($"Sample item {day}");
            sheet.Cells[$"C{row}"].PutValue(day * 10);
        }
    }
}
