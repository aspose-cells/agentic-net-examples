using System;
using Aspose.Cells;

namespace AsposeCellsAddWorksheetsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Add a worksheet with a specific name using the Add(string) overload
            Worksheet firstSheet = workbook.Worksheets.Add("FirstSheet");
            // Put some sample data into the first worksheet
            firstSheet.Cells["A1"].PutValue("Hello from the first sheet!");

            // Add another worksheet with a different name
            Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");
            // Put some sample data into the second worksheet
            secondSheet.Cells["A1"].PutValue("Hello from the second sheet!");

            // Save the workbook to a file (lifecycle save)
            workbook.Save("AddedWorksheets.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook created with two worksheets and saved as AddedWorksheets.xlsx");
        }
    }
}