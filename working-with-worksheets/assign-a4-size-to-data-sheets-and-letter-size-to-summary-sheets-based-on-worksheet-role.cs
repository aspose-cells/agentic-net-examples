// Title: Set A4 paper size for data worksheets and Letter paper size for summary worksheets using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set A4 paper size for data worksheets only | How to apply Letter paper size to summary sheets with Aspose.Cells .NET | Conditional worksheet page setup based on sheet name using Aspose.Cells | Programmatically change paper size per worksheet in Aspose.Cells C# | Assign different paper sizes to worksheets by role in .NET workbook
// Tags: conditional worksheet paper size Aspose.Cells | A4 paper size for data sheets .NET | Letter paper size for summary sheets Aspose.Cells | page setup by worksheet name C# | set paper size per worksheet Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, adds data and summary worksheets, then iterates through each sheet. It sets PageSetup.PaperSize to A4 for sheets whose names do not contain "Summary" and to Letter for sheets that do, finally saving the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets (in real scenario worksheets may already exist)
        Worksheet dataSheet1 = workbook.Worksheets[workbook.Worksheets.Add()];
        dataSheet1.Name = "Data_January";

        Worksheet dataSheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
        dataSheet2.Name = "Data_February";

        Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
        summarySheet.Name = "Summary_Report";

        // Set paper size based on worksheet role
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Treat sheets whose name contains "Summary" as summary sheets
            if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Letter size for summary sheets
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }
            else
            {
                // A4 size for data sheets
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }
        }

        // Save the workbook
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
