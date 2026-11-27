// Title: How to set a worksheet tab color to teal using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet tab color to teal | Change Excel sheet tab color programmatically with Aspose.Cells .NET | C# code to apply teal tab color to a workbook worksheet using Aspose.Cells | Saving an XLSX file with a colored worksheet tab in Aspose.Cells | Set custom tab color for a worksheet in Aspose.Cells example
// Tags: Aspose.Cells worksheet tab color | C# set worksheet tab color | Excel tab color Aspose.Cells | save workbook with colored tab | XLSX tab color customization

using System;
using System.Drawing;
using Aspose.Cells;

namespace WorksheetTabColorDemo
{
    // This example creates a new Workbook, accesses the first Worksheet, assigns Color.Teal to its TabColor property, writes the color to the console, and saves the file as TabColorTeal.xlsx, demonstrating how to color a worksheet tab programmatically with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the worksheet tab color to teal
                worksheet.TabColor = Color.Teal;

                // Optional: display the set color in console
                Console.WriteLine($"Worksheet tab color set to: {worksheet.TabColor}");

                // Save the workbook to a file
                string outputPath = "TabColorTeal.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
