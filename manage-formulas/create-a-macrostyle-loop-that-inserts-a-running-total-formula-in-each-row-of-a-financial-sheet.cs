using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RunningTotalMacroStyle
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load a template if it exists)
                Workbook workbook;
                const string templatePath = "Template.xlsx";
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                }

                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header row
                cells["A1"].PutValue("Date");
                cells["B1"].PutValue("Description");
                cells["C1"].PutValue("Amount");
                cells["D1"].PutValue("Running Total");

                // Sample data (rows 2..6)
                string[] dates = { "2023-01-01", "2023-01-02", "2023-01-03", "2023-01-04", "2023-01-05" };
                string[] desc = { "Sale A", "Sale B", "Sale C", "Sale D", "Sale E" };
                double[] amounts = { 1200.50, 850.75, 430.00, 1025.25, 600.00 };

                for (int i = 0; i < dates.Length; i++)
                {
                    int row = i + 1; // Data starts at row index 1 (Excel row 2)
                    cells[row, 0].PutValue(dates[i]);   // Date
                    cells[row, 1].PutValue(desc[i]);    // Description
                    cells[row, 2].PutValue(amounts[i]); // Amount
                }

                // Insert running total formula for each data row
                // D2 = C2
                // D3 = D2 + C3, D4 = D3 + C4, etc.
                for (int i = 1; i <= dates.Length; i++) // i is zero‑based row index of data
                {
                    Cell totalCell = cells[i, 3]; // Column D (index 3)

                    if (i == 1) // First data row (Excel row 2)
                    {
                        totalCell.SetFormula("=C2", null);
                    }
                    else
                    {
                        // Example for row 3 (i = 2): =D2+C3
                        string formula = $"=D{i}+C{i + 1}";
                        totalCell.SetFormula(formula, null);
                    }
                }

                // Calculate all formulas so that values are materialized
                workbook.CalculateFormula();

                // Save the workbook
                const string outputPath = "RunningTotalDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            RunningTotalMacroStyle.Run();
        }
    }
}