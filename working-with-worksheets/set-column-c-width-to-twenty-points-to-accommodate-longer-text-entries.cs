// Title: Set Column C to 20‑character width with Aspose.Cells (C#)
// Description: The C# sample creates a workbook, accesses the first worksheet, and calls Worksheet.Cells.SetColumnWidth to make column C 20 character units wide, writes a long string to C1, and saves the file as ColumnCWidth20.xlsx.
// Keywords: Aspose.Cells C# column width | SetColumnWidth method | Excel column C width 20 characters | adjust column width Aspose.Cells .NET | worksheet column sizing | Aspose.Cells API column width | fixed column width Excel
// Common Searches: C# Aspose.Cells set column width | How to change column C width in Aspose.Cells | Set column width to 20 characters using Aspose.Cells .NET | Aspose.Cells column width example | Increase Excel column width for long text Aspose.Cells
// Developer Intent: Define a fixed 20‑character width for column C in an Excel workbook via the Aspose.Cells .NET API.
// Use Cases: Generate reports where column C holds detailed descriptions that must not wrap. | Create a template with a predefined width for comment fields in column C. | Ensure consistent layout before populating column C with variable‑length data.
// AI Prompts: Show how to set column widths for multiple columns based on their content length using Aspose.Cells for .NET. | Provide code to auto‑fit column C after inserting data with Aspose.Cells. | Explain the differences between SetColumnWidth and AutoFitColumn in the Aspose.Cells API.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The C# sample creates a workbook, accesses the first worksheet, and calls Worksheet.Cells.SetColumnWidth to make column C 20 character units wide, writes a long string to C1, and saves the file as ColumnCWidth20.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the width of column C (zero‑based index 2) to 20 character units
                worksheet.Cells.SetColumnWidth(2, 20);

                // Add sample data to demonstrate the width
                worksheet.Cells["C1"].PutValue("This is a longer text entry for column C");

                // Save the workbook
                string outputPath = "ColumnCWidth20.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
