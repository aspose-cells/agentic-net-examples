// Title: Explicitly Dispose Aspose.Cells Workbook Objects in C# to Release File Handles and Memory
// Description: This C# example demonstrates how to create a new Aspose.Cells Workbook, write data, save it, and then call Workbook.Dispose to free file handles and memory. It also shows loading the saved file, modifying a cell, saving the changes, and disposing the second workbook, illustrating proper resource cleanup for both creation and modification scenarios.
// Keywords: Aspose.Cells Workbook Dispose | C# Aspose.Cells memory management | release file handles Aspose.Cells | Aspose.Cells .NET resource cleanup | Workbook.Dispose best practice | Aspose.Cells save and close | prevent file lock Aspose.Cells | Aspose.Cells example GitHub
// Common Searches: How to dispose Aspose.Cells Workbook in C# | Release file handles after saving Aspose.Cells workbook | Aspose.Cells memory leak prevention | Workbook.Dispose usage Aspose.Cells | Aspose.Cells example for disposing workbooks
// Developer Intent: Ensure each Aspose.Cells Workbook is explicitly disposed after use to free file handles and memory, avoiding file locks and memory leaks.
// Use Cases: Create a workbook, add data, save it, and call Dispose to close the file. | Load an existing workbook, modify cells, save changes, and dispose the object. | Process a batch of workbooks sequentially, disposing each one to prevent resource exhaustion.
// AI Prompts: Generate C# code that uses a using statement to automatically dispose Aspose.Cells Workbook objects. | Refactor the provided snippet so Workbook.Dispose is called even when an exception occurs. | Explain the impact of Workbook.Dispose on file handles and memory in Aspose.Cells .NET.

using System;
using Aspose.Cells;

// This C# example demonstrates how to create a new Aspose.Cells Workbook, write data, save it, and then call Workbook.Dispose to free file handles and memory. It also shows loading the saved file, modifying a cell, saving the changes, and disposing the second workbook, illustrating proper resource cleanup for both creation and modification scenarios.
class Program
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        // Uses the Workbook() constructor rule
        Workbook workbook1 = new Workbook();

        // Access the default worksheet and add some data
        Worksheet sheet1 = workbook1.Worksheets[0];
        sheet1.Cells["A1"].PutValue("Hello Aspose.Cells!");

        // Save the workbook to disk using the Save(string) rule
        workbook1.Save("CreatedWorkbook.xlsx");

        // Explicitly release resources for the first workbook
        workbook1.Dispose();

        // ---------- Load the previously saved workbook ----------
        // Uses the Workbook(string) constructor rule for loading
        Workbook workbook2 = new Workbook("CreatedWorkbook.xlsx");

        // Modify the workbook: add current date/time to cell B2
        Worksheet sheet2 = workbook2.Worksheets[0];
        sheet2.Cells["B2"].PutValue(DateTime.Now);

        // Save the modified workbook using the Save(string, SaveFormat) rule
        workbook2.Save("ModifiedWorkbook.xlsx", SaveFormat.Xlsx);

        // Explicitly release resources for the second workbook
        workbook2.Dispose();
    }
}
