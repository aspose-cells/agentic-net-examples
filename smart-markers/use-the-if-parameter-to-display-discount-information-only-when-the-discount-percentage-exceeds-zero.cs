// Title: Apply an IF Formula with Aspose.Cells in C# to Show Discount Text Only for Positive Percentages
// Description: This example creates a workbook, adds "Discount%" and "Info" headers, writes sample values (including zero and negative), and assigns an IF formula to each row that outputs "Discount: X%" when the discount cell is greater than zero, otherwise leaves the cell empty. The file is saved as DiscountIfDemo.xlsx.
// Keywords: Aspose.Cells | C# IF formula | conditional discount display | Excel automation | dynamic formula assignment | smart markers alternative | discount percentage check
// Common Searches: Aspose.Cells IF formula C# example | show text only for positive values in Excel using Aspose | conditional discount column Aspose.Cells | C# loop to set Excel formulas with Aspose | how to hide zero discounts in generated workbook
// Developer Intent: Generate an Excel sheet that writes a discount description only when the discount rate is greater than zero.
// Use Cases: Sales reports that list discount notes solely for items with a positive discount. | Invoices where zero or negative discounts are omitted from the remarks column. | Pricing worksheets that automatically label rows with "Discount: X%" when applicable.
// AI Prompts: Create C# code using Aspose.Cells that adds an IF formula to display "Discount: X%" only for values greater than zero. | Show how to loop through rows in Aspose.Cells and set a conditional formula for discount information, then save the workbook. | Explain how to adapt the IF formula to handle null or missing discount entries while still showing text for positive percentages.

using System;
using Aspose.Cells;

// This example creates a workbook, adds "Discount%" and "Info" headers, writes sample values (including zero and negative), and assigns an IF formula to each row that outputs "Discount: X%" when the discount cell is greater than zero, otherwise leaves the cell empty. The file is saved as DiscountIfDemo.xlsx.
class DiscountIfDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Discount%");
        sheet.Cells["B1"].PutValue("Info");

        // Sample discount percentages (0, 15%, 5%, -2%)
        sheet.Cells["A2"].PutValue(0.0);
        sheet.Cells["A3"].PutValue(0.15);
        sheet.Cells["A4"].PutValue(0.05);
        sheet.Cells["A5"].PutValue(-0.02);

        // Apply IF formula: show discount text only when the percentage is greater than zero
        for (int row = 2; row <= 5; row++)
        {
            string discountCell = $"A{row}";
            string infoCell = $"B{row}";
            // Formula: IF(Ax>0, "Discount: " & TEXT(Ax,"0%"), "")
            string formula = $"IF({discountCell}>0, \"Discount: \" & TEXT({discountCell}, \"0%\"), \"\")";
            sheet.Cells[infoCell].Formula = formula;
        }

        // Save the workbook
        workbook.Save("DiscountIfDemo.xlsx");
    }
}
