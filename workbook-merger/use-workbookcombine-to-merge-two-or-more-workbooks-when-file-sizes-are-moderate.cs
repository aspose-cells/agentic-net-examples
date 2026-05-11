using System;
using Aspose.Cells;

namespace AsposeCellsCombineDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create the first source workbook and put some data in it
            Workbook sourceWorkbook1 = new Workbook();
            sourceWorkbook1.Worksheets[0].Cells["A1"].PutValue("Data from Source Workbook 1");

            // Create the second source workbook and put some data in it
            Workbook sourceWorkbook2 = new Workbook();
            sourceWorkbook2.Worksheets[0].Cells["A1"].PutValue("Data from Source Workbook 2");

            // Create the destination workbook that will receive the combined data
            Workbook destinationWorkbook = new Workbook();

            // Combine the first source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook1);

            // Combine the second source workbook into the destination workbook
            destinationWorkbook.Combine(sourceWorkbook2);

            // Save the combined workbook to disk in XLSX format
            destinationWorkbook.Save("CombinedWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}