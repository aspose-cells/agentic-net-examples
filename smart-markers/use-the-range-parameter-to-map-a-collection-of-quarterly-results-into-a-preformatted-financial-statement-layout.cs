// Title: Map a collection of quarterly results into a pre‑formatted financial statement using Aspose.Cells Range in C#
// Common Searches: how to insert a list of objects into a specific range of an Excel template with Aspose.Cells C# | using Aspose.Cells CreateRange to populate financial statement rows from a collection | apply currency format to numeric columns in an Aspose.Cells range C# | populate Excel financial report template with quarterly data using Aspose.Cells | C# Aspose.Cells map collection to worksheet cells starting at B5
// Tags: Aspose.Cells CreateRange method C# | populate Excel template from collection Aspose.Cells | apply currency style to range cells Aspose.Cells | financial statement generation Aspose.Cells | map quarterly results to worksheet range C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a pre‑formatted financial statement workbook, defines the top‑left cell (B5) for data insertion, creates a Range sized to the quarterly results collection, writes quarter, revenue, expense, and profit values into the range, applies a built‑in currency style to the numeric columns, and saves the populated workbook as a new file.
class QuarterlyResult
{
    public string Quarter { get; set; } = string.Empty;
    public double Revenue { get; set; }
    public double Expense { get; set; }
    public double Profit { get; set; }
}

class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "FinancialStatementTemplate.xlsx";
            const string resultPath = "FinancialStatementResult.xlsx";

            // Ensure the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the pre‑formatted financial statement template
            Workbook workbook = new Workbook(templatePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Sample collection of quarterly results
            List<QuarterlyResult> results = new List<QuarterlyResult>
            {
                new QuarterlyResult { Quarter = "Q1", Revenue = 120000, Expense = 80000, Profit = 40000 },
                new QuarterlyResult { Quarter = "Q2", Revenue = 150000, Expense = 90000, Profit = 60000 },
                new QuarterlyResult { Quarter = "Q3", Revenue = 130000, Expense = 85000, Profit = 45000 },
                new QuarterlyResult { Quarter = "Q4", Revenue = 160000, Expense = 95000, Profit = 65000 }
            };

            // Define the top‑left cell of the data block in the template (e.g., B5)
            int startRow = 4;      // Row index (zero‑based) for row 5
            int startColumn = 1;   // Column index (zero‑based) for column B

            // Create a range sized to the data (rows = results.Count, columns = 4)
            Aspose.Cells.Range dataRange = sheet.Cells.CreateRange(startRow, startColumn, results.Count, 4);

            // Populate the range with the quarterly data
            for (int i = 0; i < results.Count; i++)
            {
                dataRange[i, 0].PutValue(results[i].Quarter);
                dataRange[i, 1].PutValue(results[i].Revenue);
                dataRange[i, 2].PutValue(results[i].Expense);
                dataRange[i, 3].PutValue(results[i].Profit);
            }

            // Apply currency format to the numeric columns (Revenue, Expense, Profit)
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 10; // Built‑in currency format

            // Apply the style to numeric cells individually (columns 1‑3)
            for (int i = 0; i < results.Count; i++)
            {
                for (int col = 1; col <= 3; col++)
                {
                    dataRange[i, col].SetStyle(currencyStyle);
                }
            }

            // Save the populated financial statement
            workbook.Save(resultPath);
            Console.WriteLine($"Financial statement saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
