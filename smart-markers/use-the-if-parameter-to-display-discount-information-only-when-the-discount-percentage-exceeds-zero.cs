using System;
using Aspose.Cells;

class DiscountDisplayDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Price");
        sheet.Cells["B1"].PutValue("Discount%");
        sheet.Cells["C1"].PutValue("DiscountInfo");

        // Sample data rows
        sheet.Cells["A2"].PutValue(100);
        sheet.Cells["B2"].PutValue(0.15); // 15% discount

        sheet.Cells["A3"].PutValue(200);
        sheet.Cells["B3"].PutValue(0);    // No discount

        sheet.Cells["A4"].PutValue(150);
        sheet.Cells["B4"].PutValue(0.05); // 5% discount

        // Use IF formula to display discount information only when discount > 0
        // The formula returns a text like "Discount: 15%" or an empty string
        sheet.Cells["C2"].Formula = "IF(B2>0, \"Discount: \" & TEXT(B2, \"0%\"), \"\")";
        sheet.Cells["C3"].Formula = "IF(B3>0, \"Discount: \" & TEXT(B3, \"0%\"), \"\")";
        sheet.Cells["C4"].Formula = "IF(B4>0, \"Discount: \" & TEXT(B4, \"0%\"), \"\")";

        // Evaluate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("DiscountInfo.xlsx");
    }
}