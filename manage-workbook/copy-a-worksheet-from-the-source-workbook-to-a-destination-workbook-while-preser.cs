using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook from a file
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0]; // first worksheet

        // Create a new (empty) destination workbook
        Workbook destWorkbook = new Workbook();
        // Remove the default sheet so we can add a copy with the same name
        destWorkbook.Worksheets.Clear();

        // Add a new worksheet to the destination workbook
        Worksheet destSheet = destWorkbook.Worksheets.Add(sourceSheet.Name);

        // Copy the source worksheet's contents and formats (styles) to the destination worksheet
        destSheet.Copy(sourceSheet);

        // Save the destination workbook to a file
        destWorkbook.Save("destination.xlsx");
    }
}