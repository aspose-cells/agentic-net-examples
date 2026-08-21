// Title: Aspose.Cells .NET – Update ListObject Table Comment with Version and Author Initials (C#)
// Description: Demonstrates how to create a workbook with Aspose.Cells for .NET, add a ListObject (Excel table), and set its Comment property to a custom string that includes version information and author initials. The example saves the workbook as UpdatedTableComment.xlsx and prints the comment for verification.
// Keywords: Aspose.Cells C# update table comment | ListObject comment .NET | Excel table metadata Aspose.Cells | add version to table comment | author initials in Excel table | Aspose.Cells ListObject Comment property | C# Aspose.Cells example
// Common Searches: how to change ListObject comment using Aspose.Cells for .NET | add version number to Excel table comment C# | Aspose.Cells set author initials in table comment | update Aspose.Cells table comment after creation | store revision info in Excel table with Aspose.Cells
// Developer Intent: Set a ListObject's Comment to include version and author initials for documentation tracking.
// Use Cases: Embed revision metadata directly in an Excel table by updating ListObject.Comment. | Verify the comment value at runtime by printing it to the console before saving. | Persist the comment in the .xlsx file so downstream processes can read version and author information.
// AI Prompts: Generate C# code that creates a workbook, adds a ListObject, and assigns a comment containing a version number and author initials using Aspose.Cells. | Explain where the ListObject.Comment value is stored in the .xlsx package and how to retrieve it later with Aspose.Cells. | Provide best‑practice error handling for updating a table comment in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with Aspose.Cells for .NET, add a ListObject (Excel table), and set its Comment property to a custom string that includes version information and author initials. The example saves the workbook as UpdatedTableComment.xlsx and prints the comment for verification.
    public class UpdateTableCommentDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (A1:B3)
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue(200);

                // Add a ListObject (table) covering the data range
                int listObjectIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
                ListObject listObject = worksheet.ListObjects[listObjectIndex];

                // Update the table comment to include version info and author initials
                string versionInfo = "Version 1.2";
                string authorInitials = "JD";
                listObject.Comment = $"{versionInfo} - Author: {authorInitials}";

                // Optional: display the comment in console for verification
                Console.WriteLine("Updated Table Comment: " + listObject.Comment);

                // Save the workbook
                workbook.Save("UpdatedTableComment.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as UpdatedTableComment.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UpdateTableCommentDemo.Run();
        }
    }
}
