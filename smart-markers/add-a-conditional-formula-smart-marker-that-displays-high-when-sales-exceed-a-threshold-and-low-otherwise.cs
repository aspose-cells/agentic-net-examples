// Title: Add an IF conditional smart marker in Aspose.Cells for .NET to label sales as High or Low
// Common Searches: how to use IF smart marker in Aspose.Cells C# to set status based on sales value | Aspose.Cells C# example inserting conditional formula for High/Low sales column | create a conditional smart marker that returns "High" when cell value > 1000 in .NET | populate Excel with smart markers and conditional IF formula using Aspose.Cells | C# code to add conditional status column with smart markers in an Aspose.Cells workbook
// Tags: conditional smart marker IF Aspose.Cells | sales high low smart marker C# | insert Excel formula with Aspose.Cells | generate conditional status column .NET | save workbook with formulas Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, adds "Sales" and "Status" headers, fills column A with sample sales numbers, applies an IF formula as a smart marker in column B that outputs "High" when the sales value exceeds 1000 and "Low" otherwise, and saves the file as ConditionalSmartMarker.xlsx.
class ConditionalSmartMarkerExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set headers for Sales and Status columns
            sheet.Cells["A1"].PutValue("Sales");
            sheet.Cells["B1"].PutValue("Status");

            // Sample sales data
            double[] salesData = { 800, 1200, 950, 1500 };
            for (int i = 0; i < salesData.Length; i++)
            {
                // Populate sales values in column A (starting from row 2)
                sheet.Cells[i + 1, 0].PutValue(salesData[i]);

                // Insert conditional formula in column B for each row
                // Displays "High" if sales > 1000, otherwise "Low"
                string formula = $"=IF(A{i + 2}>1000,\"High\",\"Low\")";
                sheet.Cells[i + 1, 1].Formula = formula;
            }

            // Save the workbook to a file
            string outputPath = "ConditionalSmartMarker.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
