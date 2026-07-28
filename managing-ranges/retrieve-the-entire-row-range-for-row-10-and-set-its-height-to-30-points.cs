// Title: Aspose.Cells for .NET – Set Row 10 Height to 30 Points (C#)
// Description: Demonstrates how to create a workbook, select the entire 10th row using Cells.CreateRange("10:10"), set its RowHeight to 30 points, verify the value, and save the file as RowHeightDemo.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel row height | SetRowHeight | CreateRange | 30 points | worksheet row formatting | Aspose.Cells API | Excel automation .NET | row height property
// Common Searches: Aspose.Cells set row height C# | How to change height of a specific row in Excel using Aspose.Cells | CreateRange row height example Aspose.Cells .NET | Set row 10 height to 30 points Aspose.Cells | Excel row formatting with Aspose.Cells for .NET
// Developer Intent: Set the height of row 10 to 30 points in an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Standardize header row height for printable reports. | Ensure uniform row size before bulk data export. | Programmatically adjust row height to accommodate wrapped text.
// AI Prompts: Write C# code with Aspose.Cells that sets row 15 height to 25 points and saves the workbook. | Show how to retrieve a row range and conditionally change its height based on cell content using Aspose.Cells for .NET. | Provide an example that reads back the RowHeight after setting it with Cells.CreateRange to confirm the change.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, select the entire 10th row using Cells.CreateRange("10:10"), set its RowHeight to 30 points, verify the value, and save the file as RowHeightDemo.xlsx with Aspose.Cells for .NET.
    public class SetRowHeightDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the entire row range for row 10 (1‑based index) and set its height to 30 points
                // "10:10" specifies the 10th row; RowHeight works in points
                worksheet.Cells.CreateRange("10:10").RowHeight = 30;

                // Optionally, verify the height
                double height = worksheet.Cells.CreateRange("10:10").RowHeight;
                Console.WriteLine($"Row 10 height set to: {height} points");

                // Save the workbook
                workbook.Save("RowHeightDemo.xlsx");
                Console.WriteLine("Workbook saved as RowHeightDemo.xlsx");
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
            SetRowHeightDemo.Run();
        }
    }
}
