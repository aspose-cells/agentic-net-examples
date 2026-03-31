using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // -------------------------------------------------
        // Worksheet 1 - "Report"
        // -------------------------------------------------
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Report";

        // Header: Left = file name, Center = sheet name, Right = current date
        sheet1.PageSetup.SetHeader(0, "&F");   // Left section
        sheet1.PageSetup.SetHeader(1, "&A");   // Center section
        sheet1.PageSetup.SetHeader(2, "&D");   // Right section

        // Footer: Left = page number, Center = custom text, Right = current time
        sheet1.PageSetup.SetFooter(0, "Page &P");   // Left section
        sheet1.PageSetup.SetFooter(1, "Confidential"); // Center section
        sheet1.PageSetup.SetFooter(2, "&T");       // Right section

        // -------------------------------------------------
        // Worksheet 2 - "Summary"
        // -------------------------------------------------
        int sheet2Index = workbook.Worksheets.Add(); // add a new sheet
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Summary";

        // Header: custom texts with font formatting
        sheet2.PageSetup.SetHeader(0, "Company XYZ"); // Left
        sheet2.PageSetup.SetHeader(1, "&C&\"Arial,Bold\"Summary Report"); // Center, bold Arial
        sheet2.PageSetup.SetHeader(2, "&R&\"Arial\"&8Page &P of &N"); // Right, page numbering

        // Footer: custom texts and date
        sheet2.PageSetup.SetFooter(0, "Prepared by: John Doe"); // Left
        sheet2.PageSetup.SetFooter(1, "&C&\"Arial\"&8&[Date]"); // Center, date
        sheet2.PageSetup.SetFooter(2, "Page &P of &N"); // Right

        // -------------------------------------------------
        // Save the workbook (lifecycle save)
        // -------------------------------------------------
        workbook.Save("DistinctHeadersFooters.xlsx");
    }
}