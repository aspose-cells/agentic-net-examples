// Title: C# – Create an Excel workbook with month‑named worksheets and sample data using Aspose.Cells
// Description: The sample program creates a new Workbook, renames the first sheet to "January", adds eleven additional worksheets named for the remaining months, inserts a header row and two rows of sample values on each sheet, and saves the file as MonthsWorkbook.xlsx. Demonstrates programmatic worksheet creation and data population with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel automation | create worksheets programmatically | month worksheets | populate Excel sheet with sample data | save workbook as .xlsx | multiple worksheets Aspose.Cells | Excel file generation .NET
// Common Searches: Aspose.Cells add worksheet for each month | C# generate Excel file with month tabs | populate multiple sheets with sample data using Aspose.Cells | rename default worksheet Aspose.Cells .NET | save workbook as .xlsx with Aspose.Cells
// Developer Intent: Generate an Excel workbook that contains twelve worksheets named after the months, each pre‑filled with a header and two rows of sample data.
// Use Cases: Build a calendar workbook where each month has its own sheet ready for data entry. | Create a template for monthly reporting that includes placeholder rows on every month tab. | Produce a test workbook for automated unit tests that requires multiple month‑named worksheets.
// AI Prompts: Write C# code with Aspose.Cells to add twelve worksheets named after the months and insert a header plus two sample rows on each sheet. | Extend the example to add a formula that sums the numeric values on every month worksheet. | Modify the program to write the workbook to a MemoryStream instead of a physical file.

using System;
using Aspose.Cells;

// The sample program creates a new Workbook, renames the first sheet to "January", adds eleven additional worksheets named for the remaining months, inserts a header row and two rows of sample values on each sheet, and saves the file as MonthsWorkbook.xlsx. Demonstrates programmatic worksheet creation and data population with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // List of month names
        string[] months = new string[]
        {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };

        // Rename the default worksheet to the first month and add sample data
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
        workbook.Save("MonthsWorkbook.xlsx");
    }

    // Helper method to insert sample data into a worksheet
    static void PopulateWorksheet(Worksheet sheet, string monthName)
    {
        // Header row
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sample Value");

        // Sample data rows
        sheet.Cells["A2"].PutValue(monthName);
        sheet.Cells["B2"].PutValue(100); // Example numeric value

        sheet.Cells["A3"].PutValue(monthName + " Detail");
        sheet.Cells["B3"].PutValue(200); // Another example value
    }
}
