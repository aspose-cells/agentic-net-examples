// Title: Convert an Excel workbook with an updated chart to PDF while preserving chart formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads a .xlsx file, optionally changes the first chart's title, and saves the workbook as a PDF with chart formatting intact using Aspose.Cells. | Show how to verify the source Excel file exists and implement exception handling when converting a workbook that contains charts to PDF in C# with Aspose.Cells. | Provide the minimal Aspose.Cells API calls needed to export an entire workbook, including embedded charts, to a PDF file.
// Common Searches: how to export an Excel file with charts to PDF using Aspose.Cells C# | preserve chart appearance when converting XLSX to PDF in .NET | Aspose.Cells C# example for saving workbook as PDF with chart formatting | convert workbook to PDF and keep chart title formatting Aspose.Cells | C# code to check file existence before Aspose.Cells PDF export
// Tags: Aspose.Cells workbook.Save PDF with charts | preserve chart formatting Aspose.Cells C# | check Excel file existence Aspose.Cells | exception handling Aspose.Cells PDF conversion | update chart title before PDF export Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample checks for 'input.xlsx', loads it into an Aspose.Cells Workbook, optionally updates the first chart's title, and then saves the entire workbook as 'output.pdf' using SaveFormat.Pdf, which automatically retains chart formatting while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Access the first chart on the worksheet (adjust index if needed)
            if (sheet.Charts.Count > 0)
            {
                Chart chart = sheet.Charts[0];

                // Example of updating the chart (optional)
                // chart.Title.Text = "Updated Chart Title";
                // chart.Title.IsVisible = true;
            }

            // Export the entire workbook, including the chart, to a PDF file
            // The chart formatting is preserved automatically by Aspose.Cells
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
