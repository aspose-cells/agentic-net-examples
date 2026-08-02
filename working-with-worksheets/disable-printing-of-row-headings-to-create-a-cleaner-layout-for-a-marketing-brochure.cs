// Title: Hide Row and Column Headings When Exporting a Worksheet to PDF with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, populate product data, turn off PrintHeadings, set a custom PrintArea, and save the sheet as a PDF ideal for a clean marketing brochure.
// Keywords: Aspose.Cells hide headings PDF | C# export worksheet without row numbers | PrintHeadings false Aspose.Cells | set print area Aspose.Cells PDF | remove column letters PDF export .NET
// Common Searches: how to hide row and column headings in Aspose.Cells PDF export | Aspose.Cells .NET disable print headings for brochure | set print area and turn off headings when saving as PDF | export Excel sheet to PDF without headings using Aspose
// Developer Intent: Prevent row numbers and column letters from appearing in the PDF output so the brochure shows only the table data.
// Use Cases: Generate a product catalog PDF with a tidy layout, free of worksheet headings. | Create a sales summary brochure where only the data grid is visible. | Export pricing tables for client presentations without Excel row/column labels.
// AI Prompts: Write C# code with Aspose.Cells that disables row and column headings, defines a print area, and saves the worksheet as a PDF. | Show how to configure PageSetup.PrintHeadings and PageSetup.PrintArea before exporting to PDF for a marketing brochure. | Explain the steps to produce a PDF brochure from a worksheet using Aspose.Cells while omitting worksheet headings.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, populate product data, turn off PrintHeadings, set a custom PrintArea, and save the sheet as a PDF ideal for a clean marketing brochure.
    public class DisableRowHeadingsPrintDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data that will appear in the brochure
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["B3"].PutValue(1.8);

                // Disable printing of row and column headings (row numbers, column letters)
                worksheet.PageSetup.PrintHeadings = false;

                // Define the area to be printed
                worksheet.PageSetup.PrintArea = "A1:B3";

                // Save the workbook as PDF, suitable for a marketing brochure
                string outputPath = "Brochure.pdf";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableRowHeadingsPrintDemo.Run();
        }
    }
}
