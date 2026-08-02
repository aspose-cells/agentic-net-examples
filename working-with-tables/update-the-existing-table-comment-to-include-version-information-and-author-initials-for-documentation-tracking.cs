// Title: Set ListObject Table Comment with Version and Author in Aspose.Cells C#
// Description: Shows how to create a workbook, add a ListObject, and assign a comment that includes a version identifier and author initials before saving the file.
// Keywords: Aspose.Cells | C# ListObject comment | Excel table metadata | version tag | author initials | update table comment | Aspose.Cells ListObject | Excel automation | metadata tracking
// Common Searches: Aspose.Cells set table comment C# | add version info to Excel table using Aspose | store author initials in ListObject comment | update ListObject comment programmatically | Aspose.Cells comment property example
// Developer Intent: Programmatically embed version and author information into a ListObject’s comment in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Include change‑log details in generated reports so reviewers can quickly see table version and author. | Automate documentation of data tables before distributing workbooks to stakeholders. | Maintain revision history across batch‑generated Excel files by updating the comment field of each table.
// AI Prompts: Generate code that adds a timestamp to the ListObject comment together with version and author using Aspose.Cells. | Provide a C# snippet to read a ListObject comment and parse out the version and author values. | Explain how to apply the same comment update to every table in a workbook with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a ListObject, and assign a comment that includes a version identifier and author initials before saving the file.
    public class UpdateTableCommentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue(200);

                // Add a ListObject (table) covering the data range
                int listObjectIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject table = worksheet.ListObjects[listObjectIndex];

                // Update the table comment with version information and author initials
                string versionInfo = "v1.0";
                string authorInitials = "AB";
                table.Comment = $"Version: {versionInfo} | Author: {authorInitials}";

                // Optional: display the updated comment in console
                Console.WriteLine("Updated Table Comment: " + table.Comment);

                // Save the workbook
                workbook.Save("UpdatedTableComment.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                UpdateTableCommentDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unhandled exception: " + ex.Message);
            }
        }
    }
}
