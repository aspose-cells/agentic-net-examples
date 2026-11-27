// Title: How to set Author and Title built‑in properties and freeze the first row of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# asp.net set workbook author property with aspose.cells | asp.net update Excel file title built‑in document property using Aspose.Cells | freeze first row of a worksheet programmatically with Aspose.Cells C# | ensure at least two rows exist before applying FreezePanes in Aspose.Cells | save modified workbook and verify document properties with Aspose.Cells
// Tags: built-in document properties author title Aspose.Cells | freeze pane first row worksheet Aspose.Cells | update workbook metadata Excel C# | apply FreezePanes Xlsx Aspose.Cells | verify written properties Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookMetadataAndFreeze
{
    // The example loads an existing XLSX file or creates a new workbook, sets the built‑in Author and Title properties (adding them as custom properties if missing), ensures a second row exists, freezes the first row using the FreezePanes method, saves the workbook as output.xlsx, and then reloads the file to confirm that the properties were written correctly.
    class Program
    {
        static void Main(string[] args)
        {
            // Define file names
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            Workbook workbook = null;

            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputFile))
                {
                    Console.WriteLine($"Loading workbook from \"{inputFile}\"...");
                    workbook = new Workbook(inputFile);
                }
                else
                {
                    Console.WriteLine("Input file not found. Creating a new workbook.");
                    workbook = new Workbook(); // default format is Xlsx
                    // Add some sample data so the freeze pane effect can be observed
                    Worksheet wsNew = workbook.Worksheets[0];
                    wsNew.Name = "Data";
                    wsNew.Cells["A1"].PutValue("Author");
                    wsNew.Cells["B1"].PutValue("Title");
                    wsNew.Cells["A2"].PutValue("John Doe");
                    wsNew.Cells["B2"].PutValue("Sales Report");
                }

                // -----------------------------------------------------------------
                // 1. Update built‑in document properties: Author and Title
                // -----------------------------------------------------------------
                // BuiltInDocumentProperties is a collection indexed by property name.
                // The collection always contains the standard properties, so we can set them directly.
                Console.WriteLine("Updating workbook properties...");

                // Author
                if (workbook.BuiltInDocumentProperties["Author"] != null)
                {
                    workbook.BuiltInDocumentProperties["Author"].Value = "John Doe";
                }
                else
                {
                    // If for some reason the property is missing, add it as a custom property.
                    workbook.CustomDocumentProperties.Add("Author", "John Doe");
                }

                // Title
                if (workbook.BuiltInDocumentProperties["Title"] != null)
                {
                    workbook.BuiltInDocumentProperties["Title"].Value = "Quarterly Sales Report";
                }
                else
                {
                    workbook.CustomDocumentProperties.Add("Title", "Quarterly Sales Report");
                }

                // -----------------------------------------------------------------
                // 2. Freeze the first row (row 1) so that metadata stays visible while scrolling
                // -----------------------------------------------------------------
                // The FreezePanes method works with zero‑based indexes.
                // To freeze the first row we set the freeze position at row index 1 (the second row)
                // and ask for 1 frozen row, 0 frozen columns.
                Console.WriteLine("Applying freeze pane to the first row...");

                Worksheet worksheet = workbook.Worksheets[0];
                // Ensure the worksheet has at least two rows; otherwise the call would throw.
                if (worksheet.Cells.MaxDataRow < 1)
                {
                    // Add a dummy row so the freeze operation is valid.
                    worksheet.Cells[1, 0].PutValue(string.Empty);
                }

                // Freeze first row
                worksheet.FreezePanes(1, 0, 1, 0);
                // Alternative overload:
                // worksheet.FreezePanes("A2", 1, 0);

                // -----------------------------------------------------------------
                // 3. Save the modified workbook
                // -----------------------------------------------------------------
                Console.WriteLine($"Saving modified workbook to \"{outputFile}\"...");
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");

                // -----------------------------------------------------------------
                // 4. Verify that the properties were written (optional read‑back)
                // -----------------------------------------------------------------
                Workbook verification = new Workbook(outputFile);
                Console.WriteLine("Verification of written properties:");
                Console.WriteLine($"Author: {verification.BuiltInDocumentProperties["Author"]?.Value}");
                Console.WriteLine($"Title : {verification.BuiltInDocumentProperties["Title"]?.Value}");
            }
            catch (Exception ex)
            {
                // Do not hide failures – output full exception details
                Console.Error.WriteLine("An error occurred while processing the workbook:");
                Console.Error.WriteLine(ex.ToString());
                Environment.Exit(1);
            }
        }
    }
}
