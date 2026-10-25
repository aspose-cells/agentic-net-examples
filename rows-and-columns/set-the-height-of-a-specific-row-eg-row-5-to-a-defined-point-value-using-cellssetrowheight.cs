// Title: Set a specific row height in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set height of row 5 to 30 points | How to change row height in an Excel sheet with Aspose.Cells .NET | Use Cells.SetRowHeight to adjust a single row in C# | Programmatically set Excel row height by index using Aspose.Cells | Example of setting row height in a workbook with Aspose.Cells for .NET
// Tags: Aspose.Cells SetRowHeight C# | adjust Excel row height .NET | set row height points Aspose.Cells | modify worksheet row dimensions C# | row height customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example creates a workbook (or loads one), accesses the first worksheet, and uses Cells.SetRowHeight to set row 5 (zero‑based index 4) to 30 points, then saves the file as output.xlsx while handling potential exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set the height of row 5 (zero‑based index 4) to 30 points
                sheet.Cells.SetRowHeight(4, 30);

                // Save the workbook to a file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
