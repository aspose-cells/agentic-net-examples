using System;
using Aspose.Cells;

namespace AsposeCellsReplaceConcat
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace "input.xlsx" with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Replace legacy CONCATENATE function with modern CONCAT function in all formulas
            // This replaces the text "CONCATENATE(" with "CONCAT(" throughout the workbook
            workbook.Replace("CONCATENATE(", "CONCAT(");

            // Save the updated workbook (replace "output.xlsx" with desired output path)
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}