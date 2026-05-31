using System;
using System.IO;
using Aspose.Cells;

class LoadWorkbookFromStream
{
    static void Main()
    {
        // Path to the Excel file that will be loaded
        string sourcePath = "input.xlsx";

        Workbook workbook;

        // Open the file as a read‑only stream and load it into a Workbook instance
        using (FileStream stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            // The Workbook(Stream) constructor loads the workbook from the provided stream
            workbook = new Workbook(stream);
        }

        // Access the first worksheet and display the value of cell A1
        Worksheet sheet = workbook.Worksheets[0];
        Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

        // Optional: save the loaded workbook to a new file
        // workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}