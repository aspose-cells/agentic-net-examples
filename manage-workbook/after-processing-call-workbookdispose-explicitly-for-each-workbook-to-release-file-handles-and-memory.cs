using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the default worksheet and add some data
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Save the workbook to disk
        workbook.Save("output.xlsx");

        // Explicitly release unmanaged resources for this workbook
        workbook.Dispose();

        // Load the previously saved workbook
        Workbook loadedWb = new Workbook("output.xlsx");

        // Modify the loaded workbook
        Worksheet loadedWs = loadedWb.Worksheets[0];
        loadedWs.Cells["B2"].PutValue("Modified");

        // Save the modified workbook
        loadedWb.Save("output_modified.xlsx");

        // Explicitly release resources for the loaded workbook
        loadedWb.Dispose();
    }
}