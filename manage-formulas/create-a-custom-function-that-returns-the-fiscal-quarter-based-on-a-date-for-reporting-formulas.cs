// Title: Use a custom Excel formula in C# to calculate fiscal quarter in an Aspose.Cells workbook
// Common Searches: how to add a fiscal quarter formula to an Aspose.Cells worksheet using C# | C# Aspose.Cells example for quarter calculation when fiscal year starts in July | Excel formula with MONTH and INT to determine fiscal quarter in Aspose.Cells | generate a workbook that returns fiscal quarter from a date with Aspose.Cells .NET | calculate fiscal quarter in code using a custom formula in Aspose.Cells
// Tags: fiscal quarter formula Aspose.Cells C# | custom Excel formula fiscal year start July | calculate quarter from date Aspose.Cells | INT function usage in Aspose.Cells formula | save workbook with fiscal quarter calculation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes a sample date to cell A1, assigns the Excel formula =INT((MONTH(A1)-7+12)/3)+1 to cell B1 to compute the fiscal quarter for a fiscal year that begins in July, evaluates all formulas, ensures the output directory exists, and saves the file as FiscalQuarterExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Place a sample date in cell A1.
            sheet.Cells["A1"].PutValue(new DateTime(2023, 8, 15));

            // Use an Excel formula to calculate the fiscal quarter.
            // Fiscal year starts in July: Q1 = Jul‑Sep, Q2 = Oct‑Dec, Q3 = Jan‑Mar, Q4 = Apr‑Jun
            // Formula: =INT((MONTH(A1)-7+12)/3)+1
            sheet.Cells["B1"].Formula = "INT((MONTH(A1)-7+12)/3)+1";

            // Calculate all formulas.
            workbook.CalculateFormula();

            // Define output file path.
            string outputPath = "FiscalQuarterExample.xlsx";

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
