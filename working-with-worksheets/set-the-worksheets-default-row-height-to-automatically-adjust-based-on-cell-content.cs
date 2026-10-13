// Title: Auto‑fit worksheet row heights to cell content using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates or loads an Aspose.Cells workbook, fills cells with varying text lengths, invokes the method that automatically adjusts each row’s height to fit its content, and saves the workbook. | Show how to call Aspose.Cells' AutoFitRows function in a .NET project to dynamically set row heights based on the data inside the cells.
// Common Searches: Aspose.Cells C# auto fit rows after inserting multiline text | How to make Excel row height adjust automatically with Aspose.Cells .NET | C# example for AutoFitRows method in Aspose.Cells workbook | Set default row height to fit content using Aspose.Cells for .NET | Adjust Excel row height programmatically with Aspose.Cells C#
// Tags: Aspose.Cells AutoFitRows C# | auto adjust Excel row height Aspose | dynamic row height based on cell content .NET | worksheet row height auto‑fit Aspose.Cells | C# Excel row height automation Aspose

using Aspose.Cells;
using System;

// // Creates a workbook, writes short and long strings to cells A1 and A2, calls sheet.AutoFitRows() to automatically resize each row’s height to fit the cell contents, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Example data to demonstrate auto‑adjusted row height
        sheet.Cells["A1"].PutValue("Short text");
        sheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row height to increase automatically based on its content.");

        // Adjust all rows so their height fits the cell contents
        sheet.AutoFitRows();

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}
