// Title: C# – Set default row height to 20 points with Aspose.Cells
// Description: Creates a new Workbook, accesses the first Worksheet, sets worksheet.Cells.StandardHeight to 20 points (affecting rows without custom heights), adds sample text, and saves the file as DefaultRowHeightDemo.xlsx.
// Keywords: Aspose.Cells | C# set row height | StandardHeight property | default row height 20 | worksheet row height | Aspose.Cells example | Excel row height .NET
// Common Searches: Aspose.Cells set default row height .NET | C# StandardHeight property example | How to change row height for all rows in Aspose.Cells | Increase Excel row height using Aspose.Cells C#
// Developer Intent: Apply a 20‑point default height to all rows of a worksheet using Aspose.Cells for .NET.
// Use Cases: Generate reports with consistent row spacing before populating data. | Adjust existing workbooks to improve readability when exporting to PDF. | Standardize row height across multiple worksheets in a template workbook.
// AI Prompts: Write C# code that sets worksheet.Cells.StandardHeight to 20 for every sheet in a workbook and saves it as an .xlsx file. | Show how to override the default row height for specific rows after setting a global 20‑point height with Aspose.Cells. | Explain the precedence between worksheet.Cells.StandardHeight and individual row height settings in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, accesses the first Worksheet, sets worksheet.Cells.StandardHeight to 20 points (affecting rows without custom heights), adds sample text, and saves the file as DefaultRowHeightDemo.xlsx.
    public class SetDefaultRowHeightDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the default row height for the worksheet to 20 points
                // This affects all rows that do not have a custom height
                worksheet.Cells.StandardHeight = 20;

                // Optional: add some sample data to verify the height visually
                worksheet.Cells["A1"].PutValue("Row 1 with default height");
                worksheet.Cells["A2"].PutValue("Row 2 with default height");
                worksheet.Cells["A3"].PutValue("Row 3 with default height");

                // Save the workbook (lifecycle rule: save)
                workbook.Save("DefaultRowHeightDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SetDefaultRowHeightDemo.Run();
        }
    }
}
