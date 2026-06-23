using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DropDownFromLookupDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate a lookup table (e.g., A2:A6) with values
                // -------------------------------------------------
                string[] lookupValues = { "Option1", "Option2", "Option3", "Option4", "Option5" };
                for (int i = 0; i < lookupValues.Length; i++)
                {
                    // Cells are zero‑based; row index i+1 corresponds to A2, A3, ...
                    worksheet.Cells[i + 1, 0].PutValue(lookupValues[i]);
                }

                // -------------------------------------------------
                // Add data validation with an in‑cell drop‑down to O1
                // -------------------------------------------------
                // O1 is column 14 (zero‑based) and row 0
                Validation validation = worksheet.Cells[0, 14].GetValidation();

                // Set validation type to List and point to the lookup range
                validation.Type = ValidationType.List;
                // Use absolute references to the lookup range on the same sheet
                validation.Formula1 = "$A$2:$A$6";

                // Enable the in‑cell drop‑down arrow
                validation.InCellDropDown = true;

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("DropDownFromLookupDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DropDownFromLookupDemo.Run();
        }
    }
}