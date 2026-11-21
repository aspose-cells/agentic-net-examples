// Title: Set A4 paper size for data worksheets and Letter size for summary worksheets using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set A4 paper size for worksheets containing 'Data' in the name | How to assign Letter paper size to summary sheets with Aspose.Cells .NET | Change page setup per worksheet based on worksheet name using Aspose.Cells C# | Aspose.Cells example for conditional paper size by worksheet role | Set different page sizes for multiple sheets in a workbook with Aspose.Cells
// Tags: conditional worksheet page setup Aspose.Cells | set A4 paper size for data worksheets C# | assign Letter paper size to summary worksheets Aspose.Cells | page setup by worksheet name Aspose.Cells | per-sheet paper size configuration .NET

using Aspose.Cells;

// The code loads a workbook, iterates through each worksheet, checks the sheet name for the keywords "Data" or "Summary", and applies A4 paper size to data sheets and Letter paper size to summary sheets via the PageSetup.PaperSize property before saving the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        var workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine role based on worksheet name
            // If the name contains "Data" treat it as a data sheet
            // If the name contains "Summary" treat it as a summary sheet
            if (sheet.Name.IndexOf("Data", System.StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Assign A4 paper size to data sheets
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
            }
            else if (sheet.Name.IndexOf("Summary", System.StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Assign Letter paper size to summary sheets
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }
            // If needed, add additional role checks here
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
