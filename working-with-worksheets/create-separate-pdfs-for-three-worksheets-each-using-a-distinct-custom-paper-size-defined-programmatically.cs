// Title: Generate separate PDF files for each worksheet with individual paper sizes using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through a workbook's worksheets, copies each one to a new workbook, sets its PageSetup.PaperSize, and saves it as an individual PDF file. | Show how to assign different PaperSizeType values (e.g., Letter, A5, A6) to separate worksheets before exporting them to PDF with Aspose.Cells. | Demonstrate creating temporary workbooks for each sheet to preserve page setup while generating PDF files named after the worksheet.
// Common Searches: asp.net aspose.cells export each worksheet to its own pdf with custom paper size | c# set different paper sizes for worksheets before saving as pdf using Aspose.Cells | how to create separate pdf files from multiple worksheets in a workbook aspose.cells | copy single worksheet to new workbook to retain page setup when converting to pdf c# | asp.net generate pdf per worksheet with varying page dimensions aspose.cells
// Tags: worksheet to pdf conversion with per-sheet paper size Aspose.Cells | set PageSetup.PaperSize for individual worksheets | temporary workbook for worksheet PDF export | export multiple worksheets as separate PDF files C# | custom paper size per worksheet Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a workbook with three worksheets, assigns a distinct PaperSizeType to each sheet, copies each worksheet into a temporary workbook to retain its page setup, and saves each temporary workbook as a separate PDF file (Worksheet_1.pdf, Worksheet_2.pdf, Worksheet_3.pdf).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();

            // Rename default sheet and add two more sheets
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate each sheet with sample data
            sheet1.Cells["A1"].PutValue("Data for Sheet 1");
            sheet2.Cells["A1"].PutValue("Data for Sheet 2");
            sheet3.Cells["A1"].PutValue("Data for Sheet 3");

            // Set paper sizes for each worksheet (using standard sizes)
            sheet1.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            sheet2.PageSetup.PaperSize = PaperSizeType.PaperA5;
            sheet3.PageSetup.PaperSize = PaperSizeType.PaperA6;

            // Export each worksheet to a separate PDF using its paper size
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                try
                {
                    // Create a temporary workbook containing only the current worksheet
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.Clear();
                    tempWb.Worksheets.AddCopy(workbook.Worksheets[i].Name);

                    string outputFile = $"Worksheet_{i + 1}.pdf";

                    // Save the temporary workbook as PDF
                    tempWb.Save(outputFile, SaveFormat.Pdf);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error exporting worksheet {i + 1}: {ex.Message}");
                }
            }

            Console.WriteLine("PDF files created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
