// Title: Apply a custom accounting number format to a named range of financial data and verify the displayed values with Aspose.Cells for .NET
// AI Prompts: Create a new workbook, populate cells A1:A5 with financial numbers, define a named range called FinancialData, apply the custom format "$#,##0.00;[Red]($#,##0.00)" to that range, and save the workbook. | Read each cell in the FinancialData range after formatting and output its displayed text to the console to confirm the custom format is applied correctly.
// Common Searches: asp.net apply custom accounting format to named range using Aspose.Cells | how to display negative currency values in red parentheses with Aspose.Cells C# | retrieve formatted cell text after applying style in Aspose.Cells | save workbook after applying custom number format to range Aspose.Cells | verify cell display after applying accounting style format in .NET
// Tags: accounting style number format Aspose.Cells | named range style application C# | StyleFlag All attribute usage Aspose.Cells | negative currency red parentheses formatting | read formatted cell text for verification

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// The example creates a workbook, fills A1:A5 with financial values, defines a named range "FinancialData", applies an accounting-style custom number format ($#,##0.00;[Red]($#,##0.00)) to the range using a Style and StyleFlag, saves the file as FinancialData.xlsx, and prints each cell's formatted display text to verify the formatting.
class ApplyCustomNumberFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate financial data in cells A1:A5
            double[] financialValues = { 1234.56, -789.01, 3456.78, -123.45, 0 };
            for (int i = 0; i < financialValues.Length; i++)
            {
                Cell cell = sheet.Cells[i, 0]; // Column A (index 0)
                cell.PutValue(financialValues[i]);
            }

            // Define a named range "FinancialData" that refers to A1:A5
            string rangeAddress = $"A1:A{financialValues.Length}";
            int nameIndex = workbook.Worksheets.Names.Add("FinancialData");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            namedRange.RefersTo = $"='{sheet.Name}'!{rangeAddress}";

            // Apply a custom number format to the range
            // Format: $#,##0.00;[Red]($#,##0.00)  (positive; negative in red with parentheses)
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "$#,##0.00;[Red]($#,##0.00)";

            // Use StyleFlag to indicate which style attributes to apply (apply all)
            StyleFlag styleFlag = new StyleFlag { All = true };

            // Get the range object and apply the style
            AsposeRange range = sheet.Cells.CreateRange(rangeAddress);
            range.ApplyStyle(customStyle, styleFlag);

            // Save the workbook
            string outputPath = "FinancialData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");

            // Verify display correctness by printing the formatted strings to console
            Console.WriteLine("Formatted values in the named range \"FinancialData\":");
            for (int row = 0; row < financialValues.Length; row++)
            {
                Cell cell = sheet.Cells[row, 0];
                // StringValue returns the cell's displayed text after applying formatting
                Console.WriteLine($"Cell {cell.Name}: {cell.StringValue}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
