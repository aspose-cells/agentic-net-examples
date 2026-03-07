using System;
using Aspose.Cells;

namespace AsposeSheetsDemo
{
    public class SheetOperations
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default first worksheet and set its name and data
            Worksheet original = workbook.Worksheets[0];
            original.Name = "Original";
            original.Cells["A1"].PutValue("Original Sheet Data");

            // Add a second worksheet and populate it
            Worksheet second = workbook.Worksheets.Add("Second");
            second.Cells["A1"].PutValue("Second Sheet Data");

            // Copy the "Original" worksheet using AddCopy by sheet name
            int copiedIndex = workbook.Worksheets.AddCopy("Original");
            Worksheet copied = workbook.Worksheets[copiedIndex];
            copied.Name = "Copied";

            // Move the copied worksheet to position 1 (second tab)
            copied.MoveTo(1);

            // Swap the first and last worksheets in the collection
            int firstIndex = 0;
            int lastIndex = workbook.Worksheets.Count - 1;
            workbook.Worksheets.SwapSheet(firstIndex, lastIndex);

            // Copy a group of worksheets using AddCopy(Worksheet[], string[])
            Worksheet[] sourceSheets = new Worksheet[] { original, second };
            string[] destNames = new string[] { "Original_Copy", "Second_Copy" };
            workbook.Worksheets.AddCopy(sourceSheets, destNames);

            // Save the workbook to a file
            workbook.Save("SheetOperationsDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SheetOperations.Run();
        }
    }
}