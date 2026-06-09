using System;
using Aspose.Cells;

namespace AsposeCellsHeaderMerge
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputFile.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge the header row cells (e.g., A1 to D1)
            // firstRow = 0, firstColumn = 0, totalRows = 1, totalColumns = 4
            cells.Merge(0, 0, 1, 4);

            // Set the font size of the merged header cell to 14
            Style headerStyle = cells[0, 0].GetStyle();
            headerStyle.Font.Size = 14;
            cells[0, 0].SetStyle(headerStyle);

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputFile.xlsx");
        }
    }
}