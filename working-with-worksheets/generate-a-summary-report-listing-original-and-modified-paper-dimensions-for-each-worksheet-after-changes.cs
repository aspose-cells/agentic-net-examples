// Title: Generate an Excel summary worksheet that lists each sheet’s original and updated paper size using Aspose.Cells for C#
// AI Prompts: Write a C# console program with Aspose.Cells that loops through every worksheet, records the current PageSetup.PaperSize, changes it to PaperA4, and adds a new worksheet named "Summary" containing the sheet name, original size, and new size. | Enhance the program to also read the PageSetup.Orientation (portrait or landscape) for each worksheet and include this orientation column in the summary sheet. | Add code that writes the same summary data to a CSV file while still saving the updated Excel workbook.
// Common Searches: aspocells c# how to list original paper size of each worksheet before changing to A4 | c# Aspose.Cells create summary sheet with before and after page setup values | record page setup paper dimensions for multiple worksheets using Aspose.Cells | export worksheet paper size report to CSV with Aspose.Cells C#
// Tags: Aspose.Cells worksheet paper size report | C# capture original PageSetup.PaperSize | Aspose.Cells set worksheets to PaperA4 | create summary sheet with before after page setup | export Aspose.Cells summary to CSV

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace PaperDimensionReport
{
    // Creates a workbook with three worksheets, records each sheet's initial PageSetup.PaperSize, changes all sheets to A4, adds a "Summary" worksheet that lists the sheet name, original size, and modified size, auto‑fits the columns, and saves the file as PaperDimensionReport.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook with three worksheets for demonstration
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // List to hold original and modified paper sizes for each worksheet
                var reportData = new List<(string SheetName, string OriginalSize, string ModifiedSize)>();

                // Iterate through each worksheet, capture original size, modify it, then capture new size
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Skip the summary sheet if it already exists
                    if (sheet.Name.Equals("Summary", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Capture original paper size as string
                    string originalSize = sheet.PageSetup.PaperSize.ToString();

                    // Example modification: set all sheets to A4 size
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

                    // Capture modified paper size as string
                    string modifiedSize = sheet.PageSetup.PaperSize.ToString();

                    // Store data for the report
                    reportData.Add((sheet.Name, originalSize, modifiedSize));
                }

                // Add a new worksheet for the summary report
                int summaryIndex = workbook.Worksheets.Add();
                Worksheet summarySheet = workbook.Worksheets[summaryIndex];
                summarySheet.Name = "Summary";

                // Write headers
                Cells cells = summarySheet.Cells;
                cells["A1"].PutValue("Worksheet");
                cells["B1"].PutValue("Original Paper Size");
                cells["C1"].PutValue("Modified Paper Size");

                // Populate report rows
                int rowIndex = 1; // zero‑based index; row 1 is the second row (A2, B2, C2)
                foreach (var entry in reportData)
                {
                    cells[rowIndex, 0].PutValue(entry.SheetName);
                    cells[rowIndex, 1].PutValue(entry.OriginalSize);
                    cells[rowIndex, 2].PutValue(entry.ModifiedSize);
                    rowIndex++;
                }

                // Auto‑fit columns for better readability
                summarySheet.AutoFitColumns();

                // Save the workbook
                string outputPath = "PaperDimensionReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Report saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
