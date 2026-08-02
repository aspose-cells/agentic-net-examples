// Title: Set A4 for Data worksheets and Letter for Summary worksheets using Aspose.Cells for .NET
// Description: Creates a workbook, adds a data sheet and a summary sheet, then loops through all worksheets and assigns PaperSize.A4 to sheets whose name contains "Data" and PaperSize.Letter to sheets whose name contains "Summary" before saving the file as an XLSX document.
// Keywords: Aspose.Cells | C# | .NET | worksheet paper size | A4 data sheet | Letter summary sheet | PageSetup PaperSize | conditional sheet formatting | Excel printing layout
// Common Searches: Aspose.Cells set paper size by worksheet name | C# assign A4 to data tabs and Letter to summary tabs | How to change page setup paper size for specific sheets in Aspose.Cells | Conditional paper size example Aspose.Cells .NET | Set different print sizes for Excel worksheets programmatically
// Developer Intent: Programmatically apply distinct paper sizes to worksheets based on their role (e.g., data vs. summary) within an Excel workbook.
// Use Cases: Produce printable workbooks where detailed data sheets use A4 and executive summaries use Letter to match corporate printing standards. | Automate multi‑section report generation that requires different page formats for each section without manual intervention. | Prepare Excel files for batch printing, ensuring each worksheet prints on the appropriate paper size according to its content type.
// AI Prompts: Generate C# code that iterates through an Aspose.Cells workbook and sets PaperSize.A4 for worksheets with "Data" in the name and PaperSize.Letter for those with "Summary". | Extend the role‑based paper size logic to add a "Chart" worksheet that should use PaperSize.A3. | Explain how to read, modify, and save the PageSetup.PaperSize property for each worksheet in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds a data sheet and a summary sheet, then loops through all worksheets and assigns PaperSize.A4 to sheets whose name contains "Data" and PaperSize.Letter to sheets whose name contains "Summary" before saving the file as an XLSX document.
class RoleBasedPaperSize
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add example worksheets with role-indicative names
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "DataSheet1";

        Worksheet summarySheet = workbook.Worksheets.Add("SummarySheet1");

        // Assign paper sizes based on worksheet role inferred from its name
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Data sheets → A4
            if (sheet.Name.IndexOf("Data", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }
            // Summary sheets → Letter
            else if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }
        }

        // Save the workbook
        workbook.Save("RoleBasedPaperSize.xlsx", SaveFormat.Xlsx);
    }
}
