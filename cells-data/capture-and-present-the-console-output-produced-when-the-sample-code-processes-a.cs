using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConsoleCapture
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook and save it to disk (lifecycle: create & save)
            // -----------------------------------------------------------------
            string filePath = "sample.xlsx";

            // Create a new workbook (uses Workbook() constructor)
            Workbook createWorkbook = new Workbook();

            // Access the first worksheet and add some data
            Worksheet ws = createWorkbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Hello");
            ws.Cells["B1"].PutValue("World");

            // Save the workbook to a file (uses Workbook.Save(string) method)
            createWorkbook.Save(filePath);
            createWorkbook.Dispose();

            // -----------------------------------------------------------------
            // 2. Prepare to capture console output
            // -----------------------------------------------------------------
            var stringWriter = new StringWriter();
            TextWriter originalOut = Console.Out;
            Console.SetOut(stringWriter);

            // -----------------------------------------------------------------
            // 3. Load the workbook from the file and write information to console
            //    (lifecycle: load)
            // -----------------------------------------------------------------
            Workbook loadWorkbook = new Workbook(filePath); // uses Workbook(string) constructor

            Worksheet loadedWs = loadWorkbook.Worksheets[0];

            // Write some details about the loaded workbook
            Console.WriteLine("Worksheet Name: " + loadedWs.Name);
            Console.WriteLine("Number of Cells: " + loadedWs.Cells.Count);
            Console.WriteLine("Cell A1 Value: " + loadedWs.Cells["A1"].StringValue);
            Console.WriteLine("Cell B1 Value: " + loadedWs.Cells["B1"].StringValue);

            loadWorkbook.Dispose();

            // -----------------------------------------------------------------
            // 4. Restore original console output and display the captured text
            // -----------------------------------------------------------------
            Console.SetOut(originalOut);
            string captured = stringWriter.ToString();

            Console.WriteLine("=== Captured Console Output ===");
            Console.WriteLine(captured);
        }
    }
}