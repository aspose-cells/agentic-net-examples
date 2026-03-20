using System;
using Aspose.Cells;

namespace AsposeCellsLoadSxcExample
{
    class Program
    {
        static void Main()
        {
            // Path to the StarOffice Calc (.sxc) file
            string sxcPath = "sample.sxc";

            // Create LoadOptions specifying the SXC format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Sxc);

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sxcPath, loadOptions);

            // Example: access the first worksheet and read a cell value
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First worksheet name: " + sheet.Name);
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // (Optional) modify the workbook – add a value to cell B2
            sheet.Cells["B2"].PutValue("Modified");

            // (Optional) save the modified workbook to a new file
            workbook.Save("modified.xlsx", SaveFormat.Xlsx);
        }
    }
}