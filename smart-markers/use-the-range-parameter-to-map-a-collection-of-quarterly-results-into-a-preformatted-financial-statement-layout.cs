// Title: Map a List of quarterly results into a pre‑formatted financial statement Excel template using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an Excel template with Aspose.Cells, defines a range starting at A2, and writes the entire List<QuarterResult> (Quarter, Value) into that range in a single call. | Add validation to the mapping method that checks the template contains enough rows for the quarter collection and throws a clear exception when the capacity is insufficient. | Enhance the example to auto‑size the columns and apply currency formatting to the Value column after populating the financial statement.
// Common Searches: aspnet cells fill template with quarterly results c# | c# Aspose.Cells insert list data into predefined Excel layout | using Aspose.Cells range to map collection to cells starting at A2 | check Excel template capacity before writing data Aspose.Cells
// Tags: Aspose.Cells create range for collection | write quarterly results to Excel template C# | validate template row capacity Aspose.Cells | apply currency format to Excel column Aspose.Cells | save populated workbook Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads a financial statement workbook, iterates over a List<QuarterResult>, and writes each quarter and its value into columns A and B starting at row 2. It validates the template size, auto‑sizes columns, formats values as currency, and saves the populated workbook to the specified output path.
public class QuarterResult
{
    public string Quarter { get; set; } = string.Empty;
    public decimal Value { get; set; }
}

public class FinancialStatementMapper
{
    public void MapQuarterlyResults(List<QuarterResult> results, string templatePath, string outputPath)
    {
        // Verify template file existence
        if (!File.Exists(templatePath))
            throw new FileNotFoundException($"Template file not found: {templatePath}");

        try
        {
            // Load the template workbook
            Workbook workbook = new Workbook(templatePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Starting cell (row 2, column A) – zero‑based indices
            int startRow = 1;
            int startColumn = 0;

            // Populate cells row by row
            for (int i = 0; i < results.Count; i++)
            {
                sheet.Cells[startRow + i, startColumn].PutValue(results[i].Quarter);
                sheet.Cells[startRow + i, startColumn + 1].PutValue(results[i].Value);
            }

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the populated workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error processing Excel file: {ex.Message}");
            throw;
        }
    }
}

public static class Program
{
    public static void Main()
    {
        // Sample data
        var results = new List<QuarterResult>
        {
            new QuarterResult { Quarter = "Q1 2024", Value = 12500.75m },
            new QuarterResult { Quarter = "Q2 2024", Value = 13800.00m },
            new QuarterResult { Quarter = "Q3 2024", Value = 14250.30m },
            new QuarterResult { Quarter = "Q4 2024", Value = 15500.10m }
        };

        // Paths (adjust as needed)
        string templatePath = "FinancialTemplate.xlsx";
        string outputPath   = "FinancialStatement_Output.xlsx";

        var mapper = new FinancialStatementMapper();

        try
        {
            mapper.MapQuarterlyResults(results, templatePath, outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
