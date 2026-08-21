// Title: Validate Named Range Addresses in A1 Style with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds an A1‑style and an R1C1‑style named range, then uses ConvertFormulaReferenceStyle and GetRefersTo to detect and report ranges that are not expressed in A1 notation before saving the file.
// Keywords: Aspose.Cells | C# | .NET | named range validation | A1 reference style | R1C1 to A1 conversion | ConvertFormulaReferenceStyle | GetRefersTo | workbook reference format | Excel automation
// Common Searches: how to check if a named range uses A1 notation in Aspose.Cells | convert R1C1 named range to A1 with Aspose.Cells .NET | Aspose.Cells GetRefersTo vs RefersTo difference | detect mixed reference styles in Excel workbook using C# | validate named ranges before exporting with Aspose.Cells
// Developer Intent: Identify and confirm that every named range in a workbook is defined using A1 reference notation.
// Use Cases: Audit generated reports to ensure compatibility with systems that require A1‑style named ranges. | Standardize mixed‑style named ranges before sharing workbooks with collaborators. | Automate the conversion of R1C1 references to A1 format during archival or migration processes.
// AI Prompts: Generate a C# method that scans Workbook.Worksheets.Names and returns names whose RefersTo is not in A1 style, using ConvertFormulaReferenceStyle for detection. | Provide code to update all R1C1 named ranges to A1 notation in an Aspose.Cells workbook and persist the changes. | Explain the effect of the GetRefersTo parameters on the returned string and how to compare it with a converted reference to spot non‑A1 styles.

using System;
using Aspose.Cells;

namespace NamedRangeA1Validation
{
    // C# example that creates a workbook, adds an A1‑style and an R1C1‑style named range, then uses ConvertFormulaReferenceStyle and GetRefersTo to detect and report ranges that are not expressed in A1 notation before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate some cells with sample data
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["B1"].PutValue(5);
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(25);

                // Create a named range using A1 reference style
                int a1Index = workbook.Worksheets.Names.Add("A1Range");
                Name a1Range = workbook.Worksheets.Names[a1Index];
                a1Range.RefersTo = "=Sheet1!$A$1:$A$3";

                // Create a named range using R1C1 reference style
                int r1c1Index = workbook.Worksheets.Names.Add("R1C1Range");
                Name r1c1Range = workbook.Worksheets.Names[r1c1Index];
                // Set the R1C1 reference directly
                r1c1Range.R1C1RefersTo = "'Sheet1'!R1C2:R3C2";

                // Validate each named range to ensure its address is in A1 style
                foreach (Name name in workbook.Worksheets.Names)
                {
                    // Original RefersTo string (may be A1 or R1C1)
                    string originalRefersTo = name.RefersTo;

                    // Convert the original reference to A1 style using the worksheet helper
                    // The conversion requires a base cell; using (0,0) which corresponds to A1
                    string convertedToA1 = sheet.ConvertFormulaReferenceStyle(originalRefersTo, false, 0, 0);

                    // Get the A1 formatted reference directly from the Name object
                    string a1Formatted = name.GetRefersTo(false, false);

                    // Determine if the original reference was already in A1 style
                    bool isA1 = string.Equals(convertedToA1, a1Formatted, StringComparison.OrdinalIgnoreCase);

                    Console.WriteLine($"Name: {name.Text}");
                    Console.WriteLine($"  Original RefersTo: {originalRefersTo}");
                    Console.WriteLine($"  Converted to A1 : {convertedToA1}");
                    Console.WriteLine($"  A1 Formatted   : {a1Formatted}");
                    Console.WriteLine($"  Is A1 style?   : {(isA1 ? "Yes" : "No")}");
                    Console.WriteLine();
                }

                // Save the workbook
                workbook.Save("NamedRangesValidation.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
