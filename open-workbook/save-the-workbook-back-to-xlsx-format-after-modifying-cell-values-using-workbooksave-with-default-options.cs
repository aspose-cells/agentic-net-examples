// Title: Modify cell A1 and save the workbook to XLSX using Aspose.Cells for .NET with default save settings
// Common Searches: Aspose.Cells C# change cell value and save workbook as XLSX | How to update a specific cell and write the file back with default options in Aspose.Cells | C# example for loading an XLSX, editing a cell, and saving without specifying save format | Saving modified workbook to the same format using Aspose.Cells Workbook.Save default | Aspose.Cells .NET update A1 cell then save workbook
// Tags: update cell value Aspose.Cells C# | Workbook.Save default options XLSX | modify worksheet cell A1 Aspose.Cells | save edited workbook to XLSX Aspose.Cells | Aspose.Cells load edit save workflow

using Aspose.Cells;
using System;

// The code loads an existing XLSX file, changes the value of cell A1 to "Hello World", and saves the workbook back to XLSX using Aspose.Cells' default save options.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Modify the value of cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Hello World");

        // Save the workbook back to XLSX format using default options
        workbook.Save("output.xlsx");
    }
}
