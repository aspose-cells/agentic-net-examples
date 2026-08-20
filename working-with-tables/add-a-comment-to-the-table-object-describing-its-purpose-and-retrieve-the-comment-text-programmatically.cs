// Title: Aspose.Cells C# – Set and Retrieve a ListObject (Table) Comment
// Description: Demonstrates how to create a workbook, add a ListObject (Excel table), assign a purpose‑describing comment via the ListObject.Comment property, read the comment back programmatically, and save the file as XLSX.
// Keywords: Aspose.Cells C# table comment | ListObject.Comment property | set Excel table description Aspose.Cells | retrieve ListObject comment .NET | Aspose.Cells add table metadata
// Common Searches: how to add a comment to a ListObject using Aspose.Cells for .NET | retrieve comment text from an Excel table with Aspose.Cells C# | Aspose.Cells example setting table description | read ListObject.Comment after saving workbook
// Developer Intent: Assign a purpose comment to a ListObject and programmatically read it back.
// Use Cases: Document the intent of a data table so downstream processes can interpret its meaning. | Extract table metadata for reporting or validation after workbook generation. | Ensure every worksheet table includes a descriptive comment before publishing.
// AI Prompts: Generate C# code with Aspose.Cells that adds a comment to each ListObject in a worksheet and prints the comments. | Explain how the ListObject.Comment property is persisted in an XLSX file and how to access it after reopening the workbook. | Create a script that iterates through all tables in a workbook, sets a unique comment for each, and logs the comments to the console.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableCommentDemo
{
    // Demonstrates how to create a workbook, add a ListObject (Excel table), assign a purpose‑describing comment via the ListObject.Comment property, read the comment back programmatically, and save the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the table (ListObject)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue(200);

            // Add a ListObject (table) covering the data range A1:B3
            int listObjectIndex = worksheet.ListObjects.Add(0, 0, 2, 1, true);
            ListObject listObject = worksheet.ListObjects[listObjectIndex];

            // Set a comment describing the purpose of the table
            listObject.Comment = "This table stores sample ID and Value pairs for demonstration.";

            // Retrieve the comment text programmatically
            string retrievedComment = listObject.Comment;
            Console.WriteLine("ListObject Comment: " + retrievedComment);

            // Save the workbook
            workbook.Save("TableCommentDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
