using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        object[] header = new object[] { "Product", "Price", "ReleaseDate" };
        cells.ImportObjectArray(header, 0, 0, false);

        object[][] rows = new object[][]
        {
            new object[] { "Apple", 1.99, new DateTime(2023, 12, 1) },
            new object[] { "Banana", 0.99, new DateTime(2023, 11, 15) },
            new object[] { "Cherry", 2.49, new DateTime(2024, 1, 5) }
        };

        for (int i = 0; i < rows.Length; i++)
        {
            cells.ImportObjectArray(rows[i], i + 1, 0, false);
        }

        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "yyyy-mm-dd";

        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        cells.ApplyColumnStyle(2, dateStyle, flag);

        workbook.Save("InMemoryArrayImport.xlsx");
    }
}