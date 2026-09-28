// Title: Update the Author built‑in document property of an Excel workbook with a contributor ID using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file with Aspose.Cells, assigns a specific contributor ID to the built‑in Author property, and saves the workbook. | Create a reusable C# method named UpdateAuthor that takes a file path and an author identifier, modifies the Author built‑in document property via Aspose.Cells, and overwrites the original file. | Show how to change the Author metadata of an Excel file and export the result to a new file while preserving all other workbook content using Aspose.Cells in C#.
// Common Searches: aspnet c# how to set author built‑in document property in an existing Excel file using Aspose.Cells | example code to update Excel workbook Author property with a contributor ID in .NET | Aspose.Cells change built‑in Author metadata without affecting other properties | C# program to modify Excel file author field and save as new file
// Tags: Aspose.Cells built‑in Author property | C# Excel metadata update | set Excel Author via Aspose.Cells | overwrite workbook after property change | contributor ID in Excel Author field

using Aspose.Cells;
using System;

// Loads 'input.xlsx', sets the built‑in Author property to the provided contributor ID, and saves the modified workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Update the built‑in Author property with the contributor identifier
        workbook.BuiltInDocumentProperties["Author"].Value = "ContributorID";

        // Save the changes to a new file (or overwrite the original)
        workbook.Save("output.xlsx");
    }
}
