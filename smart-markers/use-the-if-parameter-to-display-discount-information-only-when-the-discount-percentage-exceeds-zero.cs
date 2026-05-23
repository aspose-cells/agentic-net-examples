using Aspose.Cells;

class DiscountInfoDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Discount%");
        sheet.Cells["C1"].PutValue("DiscountInfo");

        // Sample data
        sheet.Cells["A2"].PutValue("Item1");
        sheet.Cells["B2"].PutValue(0.0);   // No discount

        sheet.Cells["A3"].PutValue("Item2");
        sheet.Cells["B3"].PutValue(0.15); // 15% discount

        sheet.Cells["A4"].PutValue("Item3");
        sheet.Cells["B4"].PutValue(0.05); // 5% discount

        // Use IF formula to display discount information only when discount > 0
        sheet.Cells["C2"].Formula = "IF(B2>0, \"Discount: \" & TEXT(B2, \"0%\"), \"\")";
        sheet.Cells["C3"].Formula = "IF(B3>0, \"Discount: \" & TEXT(B3, \"0%\"), \"\")";
        sheet.Cells["C4"].Formula = "IF(B4>0, \"Discount: \" & TEXT(B4, \"0%\"), \"\")";

        // Evaluate formulas
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("DiscountInfo.xlsx");
    }
}