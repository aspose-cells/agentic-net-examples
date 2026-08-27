// Title: How to set a uniform column width for every column in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Assign a value to worksheet.Cells.StandardWidth (e.g., 20) to give all columns the same width, then save the workbook. | After setting Cells.StandardWidth, read back the standard width and the width of a specific column to verify the change in a C# Aspose.Cells program.
// Common Searches: Aspose.Cells C# set default column width for entire worksheet | How to apply same column width to all columns in an Excel file using Aspose.Cells | Cells.StandardWidth property example in .NET | Programmatically change column width for every column with Aspose.Cells | Uniform column width Excel Aspose.Cells C# tutorial
// Tags: Aspose.Cells set uniform column width | Cells.StandardWidth property usage | default column width Excel .NET | apply column width to all columns C# | save workbook after column width change Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsUniformColumnWidth
{
    // Creates a new workbook, accesses the first worksheet, sets Cells.StandardWidth to 20 character units to give every column the same width, prints the standard and column 0 widths, and saves the file as UniformColumnWidth.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a uniform column width for all columns in the worksheet.
            // This assigns the default width that will be applied to every column.
            worksheet.Cells.StandardWidth = 20.0; // width in character units

            // Optional: verify the width of a specific column (e.g., column A)
            Console.WriteLine("Standard Width set to: " + worksheet.Cells.StandardWidth);
            Console.WriteLine("Column 0 actual width: " + worksheet.Cells.GetColumnWidth(0));

            // Save the workbook to a file
            workbook.Save("UniformColumnWidth.xlsx");
        }
    }
}
