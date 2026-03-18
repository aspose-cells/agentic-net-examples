using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (must exist)
            string inputPath = "input.xlsx";

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];

            // Read values from cells A1 and B1
            string name = sheet1.Cells["A1"].StringValue;

            double amount = 0;
            Cell amountCell = sheet1.Cells["B1"];
            if (amountCell.Type == CellValueType.IsNumeric)
            {
                amount = amountCell.DoubleValue;
            }
            else
            {
                double.TryParse(amountCell.StringValue, out amount);
            }

            Console.WriteLine($"Original data - Name: {name}, Amount: {amount}");

            // Modify the value in cell B1 (e.g., increase by 10%)
            sheet1.Cells["B1"].PutValue(amount * 1.10);

            // Add a new worksheet and populate it with summary data
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");
            summarySheet.Cells["A1"].PutValue("Metric");
            summarySheet.Cells["B1"].PutValue("Value");
            summarySheet.Cells["A2"].PutValue("Original Amount");
            summarySheet.Cells["B2"].PutValue(amount);
            summarySheet.Cells["A3"].PutValue("Adjusted Amount");
            summarySheet.Cells["B3"].PutValue(amount * 1.10);

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook has been saved to '{outputPath}'.");
        }
    }
}