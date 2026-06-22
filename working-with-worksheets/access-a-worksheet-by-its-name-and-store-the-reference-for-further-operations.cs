using System;
using Aspose.Cells;

namespace WorksheetAccessByNameDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add a new worksheet with a specific name
            Worksheet newSheet = workbook.Worksheets.Add("DataSheet");

            // Optionally, add some data to the new worksheet
            newSheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            newSheet.Cells["B2"].PutValue(DateTime.Now);

            // Access the worksheet by its name and store the reference
            Worksheet accessedSheet = workbook.Worksheets["DataSheet"];

            // Demonstrate further operations using the accessed worksheet
            // For example, write a value to cell C3
            accessedSheet.Cells["C3"].PutValue("Accessed by name");

            // Save the workbook to a file
            workbook.Save("WorksheetAccessByNameDemo.xlsx");
        }
    }
}