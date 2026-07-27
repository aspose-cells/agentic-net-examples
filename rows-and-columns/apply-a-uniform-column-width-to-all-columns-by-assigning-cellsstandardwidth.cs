using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example author

class UniformColumnWidthDemo
{
    public static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Apply a uniform column width to all columns
        worksheet.Cells.StandardWidth = 18.25; // width in characters

        // Verify the applied width
        Console.WriteLine("Standard Width set to: " + worksheet.Cells.StandardWidth);
        Console.WriteLine("First column actual width: " + worksheet.Cells.GetColumnWidth(0));

        // Save the workbook (using the provided save rule)
        workbook.Save("UniformColumnWidth.xlsx", SaveFormat.Xlsx);
    }
}