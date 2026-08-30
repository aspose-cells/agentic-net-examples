// Title: Validate #REF! errors after merging Excel workbooks and deleting a referenced sheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that merges two Excel workbooks with Aspose.Cells, removes a worksheet that is referenced by a formula, recalculates all formulas, and outputs the addresses of any cells that contain a #REF! error. | Show how to iterate through every cell in an Aspose.Cells workbook, detect cells where CellValueType.IsError equals #REF! or where the string value is "#REF!", and log the worksheet and cell name.
// Common Searches: C# Aspose.Cells find #REF! errors after merging workbooks | how to check for broken references in merged Excel file using Aspose.Cells | Aspose.Cells recalculate formulas after deleting a sheet and detect reference errors | detect #REF! error cells in a workbook with Aspose.Cells .NET | validate merged workbook formulas for missing worksheets Aspose.Cells
// Tags: Aspose.Cells C# #REF! error detection | merge Excel workbooks formula validation Aspose.Cells | recalculate formulas after worksheet deletion Aspose.Cells | iterate workbook cells for error values Aspose.Cells | validate broken references in merged workbook Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRefValidation
{
    // The program creates two source workbooks, copies their worksheets into a destination workbook, removes a sheet that is referenced by a formula to provoke a #REF! error, recalculates all formulas, scans every cell for #REF! error values or literal "#REF!" strings, reports any findings, and saves the merged workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -------------------------------------------------
                // 1. Create two source workbooks with sample data
                // -------------------------------------------------
                Workbook srcWb1 = new Workbook();
                Worksheet srcWs1 = srcWb1.Worksheets[0];
                srcWs1.Name = "DataSheet1";
                srcWs1.Cells["A1"].PutValue(10);
                srcWs1.Cells["A2"].Formula = "=A1*2"; // valid formula

                Workbook srcWb2 = new Workbook();
                Worksheet srcWs2 = srcWb2.Worksheets[0];
                srcWs2.Name = "DataSheet2";
                srcWs2.Cells["B1"].PutValue(5);
                // This formula will become invalid after we remove the referenced sheet later
                srcWs2.Cells["B2"].Formula = "='DataSheet1'!A1+10";

                // -------------------------------------------------
                // 2. Create a destination workbook and merge the source worksheets
                // -------------------------------------------------
                Workbook destWb = new Workbook(); // empty workbook

                // Copy first source worksheet
                int idx1 = destWb.Worksheets.Add(); // add empty sheet
                Worksheet destWs1 = destWb.Worksheets[idx1];
                srcWs1.Copy(destWs1);
                destWs1.Name = srcWs1.Name;

                // Copy second source worksheet
                int idx2 = destWb.Worksheets.Add();
                Worksheet destWs2 = destWb.Worksheets[idx2];
                srcWs2.Copy(destWs2);
                destWs2.Name = srcWs2.Name;

                // -------------------------------------------------
                // 3. (Optional) Simulate a situation that could cause #REF! errors
                //    For demonstration, we delete the sheet that B2 references.
                // -------------------------------------------------
                destWb.Worksheets.RemoveAt(0); // removes "DataSheet1"

                // -------------------------------------------------
                // 4. Recalculate all formulas in the merged workbook
                // -------------------------------------------------
                destWb.CalculateFormula();

                // -------------------------------------------------
                // 5. Validate that no #REF! errors exist in any cell
                // -------------------------------------------------
                bool hasRefError = false;
                foreach (Worksheet ws in destWb.Worksheets)
                {
                    Cells cells = ws.Cells;
                    // Iterate through all used cells
                    foreach (Cell cell in cells)
                    {
                        // Check for error cells. Aspose.Cells represents errors as type IsError.
                        // Since GetErrorValue may not be available in some versions, compare the string representation.
                        if (cell.Type == CellValueType.IsError && cell.Value != null && cell.Value.ToString() == "#REF!")
                        {
                            hasRefError = true;
                            Console.WriteLine($"#REF! error found at {ws.Name}!{cell.Name}");
                        }
                        // Also handle cells that contain the literal string "#REF!"
                        else if (cell.Type == CellValueType.IsString && cell.StringValue == "#REF!")
                        {
                            hasRefError = true;
                            Console.WriteLine($"#REF! error found at {ws.Name}!{cell.Name}");
                        }
                    }
                }

                if (!hasRefError)
                {
                    Console.WriteLine("No #REF! errors detected after merging workbooks.");
                }

                // -------------------------------------------------
                // 6. Save the merged workbook (optional)
                // -------------------------------------------------
                destWb.Save("MergedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
