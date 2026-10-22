// Title: Create a calculated column with IF formula to label rows as High or Low based on a numeric threshold using Aspose.Cells for .NET
// AI Prompts: Generate a new workbook, fill column A with numeric values, then add column B containing an IF formula that returns "High" when the A‑cell value exceeds a specified threshold and "Low" otherwise, and finally save the file. | Write C# code that loops through rows, assigns an IF(A{row}>threshold) formula to each cell in column B, triggers formula calculation, and outputs the workbook as CalculatedColumn.xlsx.
// Common Searches: aspnet add calculated column with IF function using Aspose.Cells | how to set IF formula for each row in Aspose.Cells C# example | categorize numeric data as High Low in Excel via Aspose.Cells .NET | apply conditional formula to a column programmatically with Aspose.Cells | Aspose.Cells calculate formulas after inserting IF statements
// Tags: Aspose.Cells conditional column generation | C# generate Excel calculated column | threshold based row labeling .NET | evaluate Excel formulas Aspose.Cells | programmatic conditional formula insertion

using Aspose.Cells;

// // Creates a workbook, writes numeric values to column A, adds a header in column B, applies an IF formula to each row to classify values above the threshold as "High" or "Low", evaluates all formulas, and saves the file as CalculatedColumn.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample numeric data in column A
        sheet.Cells["A1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(75);
        sheet.Cells["A3"].PutValue(45);
        sheet.Cells["A4"].PutValue(120);
        sheet.Cells["A5"].PutValue(30);

        // Define the numeric threshold
        double threshold = 50;

        // Header for the calculated column in column B
        sheet.Cells["B1"].PutValue("Category");

        // Apply IF formula to each data row
        for (int row = 1; row <= 5; row++) // rows 2 to 6 in Excel (0‑based index)
        {
            // Excel formula: =IF(A{row+1}>threshold,"High","Low")
            string formula = $"=IF(A{row + 1}>{threshold},\"High\",\"Low\")";
            sheet.Cells[row, 1].Formula = formula; // column index 1 = B
        }

        // Evaluate all formulas
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("CalculatedColumn.xlsx");
    }
}
