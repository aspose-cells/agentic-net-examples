using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Create a range from A1 to B3 using the CreateRange(string, string) method
            AsposeRange myRange = cells.CreateRange("A1", "B3");

            // Assign a name to the range (optional but useful for formulas)
            myRange.Name = "SampleRange";

            // Populate the range with sample data
            for (int i = 0; i < myRange.RowCount; i++)
            {
                for (int j = 0; j < myRange.ColumnCount; j++)
                {
                    // Put a string value indicating its position
                    myRange[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Save the workbook in XLSX format (lifecycle save)
            workbook.Save("CreatedRangeDemo.xlsx");
        }
    }
}