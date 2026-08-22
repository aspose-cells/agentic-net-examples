// Title: Create an Excel threshold violation report by scanning numeric cells with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to iterate through every cell in a worksheet, detect numeric values greater than a specified threshold, and collect their addresses. | Generate a new worksheet named "ThresholdReport" and fill it with two columns – the violating cell address and its numeric value – then save the workbook. | Implement a reusable method that accepts a Workbook object and a double threshold, returns a list of Cell objects whose DoubleValue exceeds the threshold, and optionally writes the results to a report sheet.
// Common Searches: asp.net c# Aspose.Cells find cells with values above 50 and export to new sheet | how to list addresses of numeric cells exceeding a threshold using Aspose.Cells | create threshold based report in Excel with Aspose.Cells .NET | scan entire worksheet for high numeric values Aspose.Cells C# example
// Tags: scan numeric cells Aspose.Cells | generate Excel threshold report C# | collect cell addresses exceeding limit Aspose.Cells | populate report worksheet Aspose.Cells | filter worksheet values by threshold .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace ThresholdReportDemo
{
    // The program creates a workbook, fills it with sample numeric data, scans every cell for values greater than a defined threshold, records the addresses of those cells, writes each address and its value to a new worksheet called "ThresholdReport", and saves the file as "ThresholdReport.xlsx" using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample numeric data (replace with your own data as needed)
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(75);
            worksheet.Cells["B1"].PutValue(55);
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["C3"].PutValue(120);
            worksheet.Cells["D4"].PutValue(45);

            // Define the numeric threshold
            double threshold = 50.0;

            // List to hold addresses of cells that exceed the threshold
            List<string> violatingAddresses = new List<string>();

            // Scan each cell in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Check if the cell contains a numeric value and exceeds the threshold
                if (cell.IsNumericValue && cell.DoubleValue > threshold)
                {
                    violatingAddresses.Add(cell.Name);
                }
            }

            // Create a new worksheet to hold the report
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "ThresholdReport";

            // Write header
            reportSheet.Cells["A1"].PutValue("Cell Address");
            reportSheet.Cells["B1"].PutValue("Value");

            // Populate the report with violating cell information
            for (int i = 0; i < violatingAddresses.Count; i++)
            {
                string address = violatingAddresses[i];
                Cell violCell = worksheet.Cells[address];
                int row = i + 2; // Start from row 2 (after header)

                reportSheet.Cells[row, 0].PutValue(address);               // Column A
                reportSheet.Cells[row, 1].PutValue(violCell.DoubleValue); // Column B
            }

            // Save the workbook with the report
            workbook.Save("ThresholdReport.xlsx");
        }
    }
}
