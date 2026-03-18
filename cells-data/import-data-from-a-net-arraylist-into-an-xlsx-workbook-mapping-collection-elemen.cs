using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsArrayListImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Prepare an ArrayList with header values
            ArrayList header = new ArrayList();
            header.Add("Name");
            header.Add("Age");
            header.Add("City");

            // Import the header horizontally starting at cell A1 (row 0, column 0)
            // isVertical = false means data will be placed in a single row
            cells.ImportArrayList(header, 0, 0, false);

            // Prepare another ArrayList with a data row
            ArrayList row1 = new ArrayList();
            row1.Add("John");
            row1.Add(30);
            row1.Add("New York");

            // Import the data row horizontally starting at cell A2 (row 1, column 0)
            cells.ImportArrayList(row1, 1, 0, false);

            // Add a second data row
            ArrayList row2 = new ArrayList();
            row2.Add("Alice");
            row2.Add(28);
            row2.Add("London");

            // Import the second data row starting at cell A3 (row 2, column 0)
            cells.ImportArrayList(row2, 2, 0, false);

            // Save the workbook to an XLSX file
            workbook.Save("ArrayListImportDemo.xlsx");
        }
    }
}