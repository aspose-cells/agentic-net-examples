// Title: Generate a weekly timeline with bold headers in Excel and export it to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create weekly timeline with bold header row and save as PDF | How to format a date column as dd-MMM-yyyy and auto‑fit columns before PDF export with Aspose.Cells | Aspose.Cells .NET generate rows for each week starting today and export worksheet to PDF | Apply bold style to header range A1:B1 using Aspose.Cells C# example | Save an Excel workbook as PDF with custom date format using Aspose.Cells for .NET
// Tags: create weekly timeline Aspose.Cells C# | apply bold style to header range Aspose.Cells | custom date format dd-MMM-yyyy Aspose.Cells | auto‑fit columns before PDF export Aspose.Cells | save workbook as PDF Aspose.Cells .NET

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, adds a bold header row, formats the Date column with the 'dd-MMM-yyyy' pattern, populates ten rows with weekly dates starting from today and placeholder events, auto‑fits the columns, and saves the worksheet as a PDF file named 'Timeline.pdf' using Aspose.Cells for .NET.
class TimelinePdfGenerator
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define header titles
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Event");

            // Apply bold font to header row using a range
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            StyleFlag headerFlag = new StyleFlag { FontBold = true };
            AsposeRange headerRange = cells.CreateRange("A1:B1");
            headerRange.ApplyStyle(headerStyle, headerFlag);

            // Set date format for the Date column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            StyleFlag dateFlag = new StyleFlag { NumberFormat = true };

            // Generate weekly intervals starting from today
            DateTime startDate = DateTime.Today;
            int numberOfWeeks = 10; // Adjust as needed

            for (int i = 0; i < numberOfWeeks; i++)
            {
                // Row index (0‑based in API). Row 0 is header, so data starts at row 1.
                int row = i + 1;

                // Calculate date for the current week
                DateTime weekDate = startDate.AddDays(i * 7);
                cells[row, 0].PutValue(weekDate);
                cells[row, 0].SetStyle(dateStyle, dateFlag); // Apply date format

                // Placeholder event description
                cells[row, 1].PutValue($"Event for week {i + 1}");
            }

            // Auto‑fit columns for better appearance
            sheet.AutoFitColumns();

            // Save the workbook as PDF
            workbook.Save("Timeline.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
