// Title: Set Workbook Author to Current User After Merging Excel Files with Aspose.Cells for .NET
// Description: This C# example demonstrates how to merge two Excel workbooks using Aspose.Cells, remove the placeholder sheet created by the default constructor, and assign the logged‑in Windows user name to both Settings.Author and BuiltInDocumentProperties.Author before saving the combined file as MergedWorkbook.xlsx.
// Keywords: Aspose.Cells | C# | .NET | workbook merge | set author property | Environment.UserName | built‑in document properties | remove default worksheet | Excel metadata
// Common Searches: Aspose.Cells set author property C# | merge multiple Excel files and set author | remove default sheet after workbook merge Aspose.Cells | Environment.UserName workbook author Aspose | update built‑in document properties after merging workbooks
// Developer Intent: Assign the current Windows user as the Author metadata of a workbook that results from merging multiple Excel files.
// Use Cases: Record the person who performed a batch merge of Excel workbooks for audit trails. | Ensure consistent author information across Excel viewers by updating both Settings.Author and BuiltInDocumentProperties.Author. | Automatically clean up the empty worksheet added by the Workbook constructor before saving the merged file.
// AI Prompts: Generate C# code using Aspose.Cells that merges two Excel workbooks, deletes the default empty worksheet, and sets the merged workbook's Author to Environment.UserName. | Show how to copy worksheets from source workbooks into a new workbook and then update Settings.Author and BuiltInDocumentProperties.Author. | Explain step‑by‑step how to modify built‑in document properties of an Aspose.Cells workbook after merging worksheets.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergeExample
{
    // This C# example demonstrates how to merge two Excel workbooks using Aspose.Cells, remove the placeholder sheet created by the default constructor, and assign the logged‑in Windows user name to both Settings.Author and BuiltInDocumentProperties.Author before saving the combined file as MergedWorkbook.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook that will hold the merged content
                Workbook mergedWorkbook = new Workbook();

                // Load the first source workbook if it exists
                Workbook source1 = null;
                if (File.Exists("Source1.xlsx"))
                {
                    source1 = new Workbook("Source1.xlsx");
                }
                else
                {
                    Console.WriteLine("Source1.xlsx not found.");
                }

                // Load the second source workbook if it exists
                Workbook source2 = null;
                if (File.Exists("Source2.xlsx"))
                {
                    source2 = new Workbook("Source2.xlsx");
                }
                else
                {
                    Console.WriteLine("Source2.xlsx not found.");
                }

                // Helper to copy worksheets from a source workbook to the merged workbook
                void CopyWorksheets(Workbook source)
                {
                    if (source == null) return;

                    foreach (Worksheet sheet in source.Worksheets)
                    {
                        // Add a new empty worksheet to the merged workbook
                        int newIndex = mergedWorkbook.Worksheets.Add();
                        Worksheet newSheet = mergedWorkbook.Worksheets[newIndex];

                        // Copy the source sheet's content into the new sheet
                        sheet.Copy(newSheet);
                    }
                }

                // Copy worksheets from both source workbooks
                CopyWorksheets(source1);
                CopyWorksheets(source2);

                // Remove the default empty worksheet that was created with the new workbook
                if (mergedWorkbook.Worksheets.Count > 0 &&
                    mergedWorkbook.Worksheets[0].Cells.MaxDataColumn == -1 &&
                    mergedWorkbook.Worksheets[0].Cells.MaxDataRow == -1)
                {
                    mergedWorkbook.Worksheets.RemoveAt(0);
                }

                // Set the author of the merged workbook to the current user name
                mergedWorkbook.Settings.Author = Environment.UserName;
                mergedWorkbook.BuiltInDocumentProperties.Author = Environment.UserName;

                // Save the merged workbook
                mergedWorkbook.Save("MergedWorkbook.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Merged workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
