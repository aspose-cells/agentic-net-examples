using System;
using Aspose.Cells;

namespace AsposeCellsPaperSizeDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample worksheets and give them role‑based names
            Worksheet dataSheet1 = workbook.Worksheets[0];
            dataSheet1.Name = "Data_January";

            Worksheet dataSheet2 = workbook.Worksheets.Add("Data_February");

            Worksheet summarySheet = workbook.Worksheets.Add("Summary_Report");

            // Iterate through all worksheets and assign paper size based on role
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // If the worksheet name contains "Data", treat it as a data sheet
                if (sheet.Name.IndexOf("Data", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Set paper size to A4
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                }
                // If the worksheet name contains "Summary", treat it as a summary sheet
                else if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // Set paper size to Letter
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
                }
                // For any other sheets, you can define a default size if needed
            }

            // Save the workbook
            workbook.Save("Workbook_With_RoleBased_PaperSizes.xlsx", SaveFormat.Xlsx);
        }
    }
}