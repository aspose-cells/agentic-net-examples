// Title: Set custom zoom (150%) and freeze the top three rows in Aspose.Cells for .NET
// Description: Creates a new Workbook, adds sample data, switches the worksheet to Normal view with a 150% zoom level, freezes the first three rows while keeping all columns scrollable using FreezePanes, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | .NET | C# | set worksheet zoom | custom view | Normal view | freeze rows | FreezePanes | Excel zoom percentage | worksheet view type | freeze top rows | Aspose.Cells example
// Common Searches: Aspose.Cells set custom zoom level | How to freeze top rows in Aspose.Cells .NET | FreezePanes method parameters Aspose.Cells | Set Normal view and 150% zoom with Aspose.Cells | C# example for freezing rows and setting zoom in Excel workbook
// Developer Intent: Create a workbook, apply a 150% zoom in Normal view, and freeze the first three rows.
// Use Cases: Generate reports where header rows stay visible while users zoom into data. | Provide spreadsheet templates that open with a predefined zoom level and frozen header rows. | Design dashboards that keep context rows fixed during horizontal scrolling.
// AI Prompts: Show how to set a 150% zoom and freeze the top three rows in Aspose.Cells for .NET with error handling. | Give an example that applies Normal view, 150% zoom, freezes the first three rows, and saves the workbook to a memory stream. | Explain the FreezePanes parameters and demonstrate freezing both rows and columns in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, adds sample data, switches the worksheet to Normal view with a 150% zoom level, freezes the first three rows while keeping all columns scrollable using FreezePanes, and saves the file as an XLSX document.
    public class CustomViewAndFreezeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some sample data (optional, just to have visible content)
                for (int i = 0; i < 20; i++)
                {
                    worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
                    worksheet.Cells[i, 1].PutValue(i * 10);
                }

                // Set the view type to Normal (default) and apply a custom zoom level (e.g., 150%)
                worksheet.ViewType = ViewType.NormalView;
                worksheet.Zoom = 150; // Zoom is a percentage between 10 and 400

                // Freeze the top 3 rows while keeping all columns scrollable
                // Parameters: row index, column index, frozen rows, frozen columns
                // Here we freeze rows at index 3 (fourth row) with 3 frozen rows and 0 frozen columns
                worksheet.FreezePanes(3, 0, 3, 0);

                // Save the workbook to a file
                string outputPath = "CustomViewAndFreezeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomViewAndFreezeDemo.Run();
        }
    }
}
