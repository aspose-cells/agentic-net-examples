using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetAccessByName
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.Name = "FirstSheet";

            // Add a second worksheet with a specific name
            Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");

            // Access the first worksheet using its name
            Worksheet accessedFirst = workbook.Worksheets["FirstSheet"];
            accessedFirst.Cells["A1"].PutValue("Accessed by name: FirstSheet");
            accessedFirst.Cells["A2"].PutValue(DateTime.Now);

            // Access the second worksheet using its name
            Worksheet accessedSecond = workbook.Worksheets["SecondSheet"];
            accessedSecond.Cells["B1"].PutValue("Accessed by name: SecondSheet");
            accessedSecond.Cells["B2"].PutValue(12345);

            // Save the workbook to a file
            workbook.Save("WorksheetAccessByName.xlsx", SaveFormat.Xlsx);
        }
    }
}