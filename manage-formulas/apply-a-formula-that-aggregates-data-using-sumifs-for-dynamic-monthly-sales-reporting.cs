// Title: Generate a monthly sales summary with SUMIFS and a month-key column using Aspose.Cells for .NET
// AI Prompts: Create C# code with Aspose.Cells that adds a helper column converting dates to "yyyy‑MM" strings and inserts a SUMIFS formula to total sales for a specified month and product. | Write a program that builds an Excel workbook, populates sample sales rows, defines input cells for target month and product, applies the SUMIFS aggregation, calculates all formulas, and saves the workbook.
// Common Searches: Aspose.Cells C# SUMIFS example with date month key | How to calculate monthly sales totals by product using Aspose.Cells | Create dynamic Excel report with SUMIFS and TEXT date conversion in .NET
// Tags: Aspose.Cells SUMIFS formula | C# Excel date to yyyy-MM column | dynamic sales aggregation .NET | criteria based SUMIFS calculation | Excel helper column TEXT function

using System;
using Aspose.Cells;

// The sample builds a new workbook, inserts sales data, creates a MonthKey helper column using the TEXT function, provides cells for target month and product, applies a SUMIFS formula to aggregate matching sales, evaluates the formulas, and saves the result as MonthlySalesReport.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // ----- Header Row -----
        cells["A1"].PutValue("Date");      // Transaction date
        cells["B1"].PutValue("Product");   // Product name
        cells["C1"].PutValue("Region");    // Sales region
        cells["D1"].PutValue("Sales");     // Sales amount
        cells["E1"].PutValue("MonthKey");  // Helper column: month in yyyy-MM format

        // ----- Sample Data (Rows 2-6) -----
        // Row 2
        cells["A2"].PutValue(new DateTime(2023, 9, 5));
        cells["B2"].PutValue("Widget");
        cells["C2"].PutValue("North");
        cells["D2"].PutValue(1200);
        // Row 3
        cells["A3"].PutValue(new DateTime(2023, 9, 12));
        cells["B3"].PutValue("Gadget");
        cells["C3"].PutValue("South");
        cells["D3"].PutValue(850);
        // Row 4
        cells["A4"].PutValue(new DateTime(2023, 9, 20));
        cells["B4"].PutValue("Widget");
        cells["C4"].PutValue("East");
        cells["D4"].PutValue(950);
        // Row 5
        cells["A5"].PutValue(new DateTime(2023, 8, 15));
        cells["B5"].PutValue("Widget");
        cells["C5"].PutValue("West");
        cells["D5"].PutValue(700);
        // Row 6
        cells["A6"].PutValue(new DateTime(2023, 9, 30));
        cells["B6"].PutValue("Widget");
        cells["C6"].PutValue("North");
        cells["D6"].PutValue(1100);

        // ----- Populate Helper Column (MonthKey) -----
        // Formula: =TEXT(A2,"yyyy-MM")
        for (int row = 2; row <= 6; row++)
        {
            cells[$"E{row}"].Formula = $"=TEXT(A{row},\"yyyy-MM\")";
        }

        // ----- Criteria Input Cells -----
        cells["G1"].PutValue("Target Month (yyyy-MM)");
        cells["G2"].PutValue(DateTime.Now.ToString("yyyy-MM")); // e.g., current month
        cells["H1"].PutValue("Target Product");
        cells["H2"].PutValue("Widget");

        // ----- Result Header -----
        cells["I1"].PutValue("Monthly Sales (SUMIFS)");

        // ----- SUMIFS Formula -----
        // Sum range: D2:D6 (Sales)
        // Criteria range 1: E2:E6 (MonthKey) = G2
        // Criteria range 2: B2:B6 (Product) = H2
        cells["I2"].Formula = "=SUMIFS(D2:D6, E2:E6, G2, B2:B6, H2)";

        // Calculate all formulas so that I2 contains the aggregated value
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("MonthlySalesReport.xlsx");
    }
}
