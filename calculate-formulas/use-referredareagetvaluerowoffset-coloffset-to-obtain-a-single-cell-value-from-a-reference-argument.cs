using System;
using Aspose.Cells;

namespace ReferredAreaGetValueDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with data
            sheet.Cells["A1"].PutValue(42);
            sheet.Cells["B1"].PutValue(100);

            // Set a formula that references cell A1
            Cell formulaCell = sheet.Cells["C1"];
            formulaCell.Formula = "=A1";

            // Ensure formulas are calculated so the reference is resolved
            workbook.CalculateFormula();

            // Retrieve the collection of referred areas (precedents) for the formula cell
            ReferredAreaCollection precedents = formulaCell.GetPrecedents();

            if (precedents != null && precedents.Count > 0)
            {
                // Get the first referred area – it should correspond to A1
                ReferredArea area = precedents[0];

                // Obtain the value at the top‑left corner of the area (offset 0,0)
                object value = area.GetValue(0, 0);

                Console.WriteLine($"Value obtained from ReferredArea: {value}");
            }
            else
            {
                Console.WriteLine("No precedents found for the formula cell.");
            }

            // Save the workbook (optional)
            workbook.Save("ReferredAreaDemo.xlsx");
        }
    }
}