// Title: Create a named range with a SUMIFS formula to aggregate monthly sales in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example to sum sales for a specific month using SUMIFS and a named range | C# code for dynamic month and year inputs to calculate monthly sales totals with Aspose.Cells | How to reference a SUMIFS result from another worksheet in an Aspose.Cells workbook | Generate a monthly sales report workbook with random data using Aspose.Cells for .NET | Using DATE and EOMONTH functions in an Aspose.Cells SUMIFS formula for monthly aggregation
// Tags: Aspose.Cells SUMIFS monthly aggregation | named range formula Aspose.Cells C# | dynamic month year parameters Excel Aspose.Cells | C# generate random sales data Aspose.Cells | Excel report sheet reference Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a new workbook, adds a 'Data' sheet with 90 days of random sales figures, and places month and year input cells. It defines a SUMIFS formula that uses DATE and EOMONTH to sum sales for the specified month, assigns the formula to a named range, and displays the result on a separate 'Report' sheet. Finally, the workbook is saved as MonthlySalesReport.xlsx.
class MonthlySalesReport
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it "Data"
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Add headers for Date and Sales columns
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Sales");

            // Generate sample sales data (90 days)
            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();
            for (int i = 0; i < 90; i++)
            {
                dataSheet.Cells[i + 1, 0].PutValue(startDate.AddDays(i));          // Date
                dataSheet.Cells[i + 1, 1].PutValue(rnd.Next(100, 1000));          // Sales
            }

            // Input cells for month and year (used by the SUMIFS formula)
            dataSheet.Cells["C1"].PutValue("Month");
            dataSheet.Cells["C2"].PutValue(2); // February
            dataSheet.Cells["D1"].PutValue("Year");
            dataSheet.Cells["D2"].PutValue(2023); // 2023

            // Define the SUMIFS formula for monthly aggregation
            string sumIfsFormula =
                "=SUMIFS(Data!$B:$B, " +
                "Data!$A:$A, \">=\"&DATE(Data!$D$2,Data!$C$2,1), " +
                "Data!$A:$A, \"<=\"&EOMONTH(DATE(Data!$D$2,Data!$C$2,1),0))";

            // Place the formula in cell E2 and label it
            dataSheet.Cells["E1"].PutValue("Monthly Sales");
            dataSheet.Cells["E2"].Formula = sumIfsFormula;

            // Create a report sheet to display the aggregated result
            Worksheet reportSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            reportSheet.Name = "Report";
            reportSheet.Cells["A1"].PutValue("Aggregated Monthly Sales:");
            // Reference the result directly from Data!E2
            reportSheet.Cells["B1"].Formula = "=Data!E2";

            // Save the workbook
            string outputPath = "MonthlySalesReport.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
