// Title: Copy a Cell Range to a New Workbook and Protect Its Structure with a Password (Aspose.Cells C#)
// Description: Demonstrates how to create a source workbook, fill range A1:B3, copy that range into a fresh workbook, apply structure protection with a password, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range | C# copy range to new workbook | protect workbook structure password | Aspose.Cells workbook protection | copy cells between workbooks | Aspose.Range example | Excel file protection C# | .NET Aspose.Cells tutorial
// Common Searches: Aspose.Cells copy range to another workbook C# | How to protect workbook structure with password using Aspose.Cells | Copy cells and lock workbook in .NET | Aspose.Cells protect workbook structure example | C# copy range and set workbook protection
// Developer Intent: Copy a defined cell range into a new workbook and secure the workbook’s structure with a password.
// Use Cases: Generate a report by extracting a summary table from a template and preventing sheet reordering. | Export a selected data slice to a separate Excel file while disallowing addition or removal of worksheets. | Create a read‑only version of a workbook by copying specific ranges and applying structure protection.
// AI Prompts: Show C# code that copies a range from one workbook to another and protects the destination workbook’s structure with a password using Aspose.Cells. | Explain how to copy multiple non‑contiguous ranges into a new workbook and apply structure protection with a custom password in Aspose.Cells for .NET. | Provide steps to verify that workbook protection was successfully applied after saving the file with Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyAndProtect
{
    // Demonstrates how to create a source workbook, fill range A1:B3, copy that range into a fresh workbook, apply structure protection with a password, and save the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Source workbook ----------
                // Create a source workbook and fill some data
                using (Workbook sourceWb = new Workbook())
                {
                    Worksheet sourceSheet = sourceWb.Worksheets[0];
                    Cells sourceCells = sourceSheet.Cells;

                    // Populate range A1:B3 with sample values
                    sourceCells["A1"].PutValue("Item");
                    sourceCells["B1"].PutValue("Quantity");
                    sourceCells["A2"].PutValue("Apple");
                    sourceCells["B2"].PutValue(10);
                    sourceCells["A3"].PutValue("Banana");
                    sourceCells["B3"].PutValue(20);

                    // Define the source range to copy (A1:B3)
                    AsposeRange sourceRange = sourceCells.CreateRange("A1:B3");

                    // ---------- Destination workbook ----------
                    // Create an empty workbook that will receive the copied range
                    using (Workbook destWb = new Workbook())
                    {
                        Worksheet destSheet = destWb.Worksheets[0];
                        Cells destCells = destSheet.Cells;

                        // Define the destination range (starting at A1, same size as source)
                        AsposeRange destRange = destCells.CreateRange("A1:B3");

                        // Copy the source range into the destination range
                        destRange.Copy(sourceRange);

                        // Protect the workbook structure with a password
                        destWb.Protect(ProtectionType.Structure, "MySecretPassword");

                        // Save the destination workbook
                        destWb.Save("CopiedAndProtectedWorkbook.xlsx", SaveFormat.Xlsx);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
