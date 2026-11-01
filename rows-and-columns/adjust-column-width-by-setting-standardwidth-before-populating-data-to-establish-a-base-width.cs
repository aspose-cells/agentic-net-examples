// Title: Set a worksheet’s StandardWidth to define a default column width before adding data with Aspose.Cells for .NET (C#)
// AI Prompts: Assign Worksheet.Cells.StandardWidth = 15 to establish a uniform column width, then fill cells and save the workbook using Aspose.Cells in C#. | Create a new Workbook, configure the first worksheet’s StandardWidth property to the desired character count, add rows of data, and export to an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells how to set default column width for all columns in C# before writing data | C# Aspose.Cells StandardWidth property example for Excel export | set global column width in Aspose.Cells workbook programmatically | initialize worksheet column width in Aspose.Cells .NET prior to populating cells
// Tags: Aspose.Cells set worksheet default column width | Worksheet.Cells.StandardWidth C# | global column width Aspose.Cells .NET | Excel export column width Aspose.Cells | initialize column width before data entry Aspose.Cells

using System;
using Aspose.Cells;

namespace AdjustColumnWidthExample
{
    // The program creates a new Workbook, sets the first worksheet's StandardWidth to 15 characters as a base column width, adds sample data to cells A1:C3, and saves the file as AdjustedColumnWidth.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set a base column width for all columns (measured in characters of the default font)
                sheet.Cells.StandardWidth = 15; // Adjust this value as needed

                // Populate sample data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");

                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(85);

                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(92);

                // Save the workbook to a file
                string outputPath = "AdjustedColumnWidth.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
