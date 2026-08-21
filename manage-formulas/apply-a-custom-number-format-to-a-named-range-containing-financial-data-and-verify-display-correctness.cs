// Title: Apply a Custom Euro Currency Number Format to a Named Range with Aspose.Cells for .NET
// Description: This C# example creates a workbook, defines a named range "FinData" (A2:A3), builds a custom Euro accounting style, applies only the number‑format part to the range using StyleFlag, reads back the Custom format from a cell to confirm the change, and saves the file as FinancialDataFormatted.xlsx.
// Keywords: Aspose.Cells | custom number format | Euro currency format | named range | StyleFlag | ApplyStyle | C# | .NET | verify cell format | financial data formatting
// Common Searches: Aspose.Cells set custom currency format for named range | How to use StyleFlag to apply only number format in Aspose.Cells | Read back custom number format after ApplyStyle Aspose.Cells | C# apply Euro accounting format to a range | Verify number format applied to cells Aspose.Cells
// Developer Intent: Apply a custom Euro accounting number format to a predefined named range and programmatically confirm that the format was applied correctly.
// Use Cases: Standardize Euro accounting display for all cells referenced by a named range without affecting other styling attributes. | Update only the number‑format of an existing range while preserving fonts, borders, and alignment. | Programmatically validate that a custom format string is set before exporting the workbook.
// AI Prompts: Generate C# code that creates a named range in Aspose.Cells, applies a custom Euro currency number format using StyleFlag, and verifies the format on a sample cell. | Explain the role of StyleFlag when applying a style to a range with Aspose.Cells and how it isolates the number‑format attribute. | Write a unit test in C# that asserts the Custom number format of a cell matches the expected Euro accounting pattern after applying a style.

using System;
using Aspose.Cells;

// This C# example creates a workbook, defines a named range "FinData" (A2:A3), builds a custom Euro accounting style, applies only the number‑format part to the range using StyleFlag, reads back the Custom format from a cell to confirm the change, and saves the file as FinancialDataFormatted.xlsx.
class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample financial data
            worksheet.Cells["A1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue(1234.56);   // Positive value
            worksheet.Cells["A3"].PutValue(-789.01);   // Negative value

            // Define a named range "FinData" that refers to the financial values
            int nameIndex = workbook.Worksheets.Names.Add("FinData");
            Name financialName = workbook.Worksheets.Names[nameIndex];
            financialName.RefersTo = "=Sheet1!$A$2:$A$3";

            // Create a style with a custom currency number format
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "_-€ * #,##0.00_-;_-€ * -#,##0.00_-;_-€ * \"-\"??_-;_-@_-";

            // Use a StyleFlag to apply only the number format part of the style
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;

            // Retrieve the range represented by the named range and apply the style
            Aspose.Cells.Range financialRange = financialName.GetRange();
            financialRange.ApplyStyle(customStyle, flag);

            // Verify that the custom format was applied to a cell in the range
            string appliedFormat = worksheet.Cells["A2"].GetStyle().Custom;
            Console.WriteLine("Applied custom format: " + appliedFormat);

            // Save the workbook
            workbook.Save("FinancialDataFormatted.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
