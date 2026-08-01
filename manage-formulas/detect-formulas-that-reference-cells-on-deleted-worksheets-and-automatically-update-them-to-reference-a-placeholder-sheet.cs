// Title: C# – Detect #REF! Formulas After Deleting a Worksheet and Redirect Them to a Placeholder Sheet with Aspose.Cells
// Description: This Aspose.Cells for .NET example creates a workbook, adds a Data and a Summary sheet, writes a formula that references the Data sheet, deletes the Data sheet, ensures a "Placeholder" sheet exists, scans every cell in the workbook, and replaces any #REF! formula with a reference to Placeholder!A1 before saving the file.
// Keywords: Aspose.Cells C# #REF! handling | update formulas after sheet deletion | placeholder worksheet Aspose.Cells | detect invalid references .NET | replace #REF! with default cell
// Common Searches: how to fix #REF! formulas after removing a worksheet using Aspose.Cells | C# code to redirect broken references to a placeholder sheet | scan all cells for #REF! in Aspose.Cells workbook | replace invalid sheet references with default value Aspose.Cells
// Developer Intent: Automatically replace formulas that become #REF! when a referenced worksheet is deleted, pointing them to a predefined placeholder sheet.
// Use Cases: Maintain functional summary reports when source data sheets are removed. | Create template workbooks that safely redirect broken references to a placeholder for later data entry. | Implement a cleanup routine that sanitizes all formulas before distributing a workbook.
// AI Prompts: Write C# code with Aspose.Cells that finds every #REF! formula after a worksheet is deleted and changes it to "Placeholder!A1". | Show how to add a placeholder worksheet if it does not exist and update all invalid formulas across a workbook using Aspose.Cells. | Explain an efficient way to iterate through all cells in an Aspose.Cells workbook to replace #REF! references with a default cell.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This Aspose.Cells for .NET example creates a workbook, adds a Data and a Summary sheet, writes a formula that references the Data sheet, deletes the Data sheet, ensures a "Placeholder" sheet exists, scans every cell in the workbook, and replaces any #REF! formula with a reference to Placeholder!A1 before saving the file.
    public class UpdateDeletedSheetReferences
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add two worksheets: "Data" and "Summary"
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";
                Worksheet summarySheet = workbook.Worksheets.Add("Summary");

                // Populate some data in the "Data" sheet
                dataSheet.Cells["A1"].PutValue(10);
                dataSheet.Cells["A2"].PutValue(20);

                // In the "Summary" sheet, add a formula that references the "Data" sheet
                // This formula will become invalid after we delete the "Data" sheet
                summarySheet.Cells["B1"].Formula = "=Data!A1+A2";

                // Ensure a placeholder sheet exists to receive updated references
                const string placeholderName = "Placeholder";
                Worksheet placeholderSheet = workbook.Worksheets[placeholderName];
                if (placeholderSheet == null)
                {
                    placeholderSheet = workbook.Worksheets.Add(placeholderName);
                    // Optionally put a default value in the placeholder cell
                    placeholderSheet.Cells["A1"].PutValue(0);
                }

                // Delete the "Data" worksheet (the one referenced by the formula)
                // After deletion, formulas that referenced it will contain #REF!
                int dataSheetIndex = workbook.Worksheets.IndexOf(dataSheet);
                workbook.Worksheets.RemoveAt(dataSheetIndex);

                // Scan all cells in all worksheets and replace any #REF! with a reference to the placeholder sheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Use the Cells iterator for efficient traversal
                    foreach (Cell cell in ws.Cells)
                    {
                        if (cell.IsFormula && cell.Formula.Contains("#REF!"))
                        {
                            // Replace the invalid reference with a reference to the placeholder sheet's A1 cell
                            cell.Formula = cell.Formula.Replace("#REF!", $"{placeholderName}!A1");
                        }
                    }
                }

                // Save the updated workbook
                workbook.Save("UpdatedReferences.xlsx");
                Console.WriteLine("Workbook saved as 'UpdatedReferences.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateDeletedSheetReferences.Run();
        }
    }
}
