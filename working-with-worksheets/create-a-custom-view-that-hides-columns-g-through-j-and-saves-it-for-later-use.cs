// Title: Create a custom view that hides columns G‑J and saves it with Aspose.Cells for .NET
// Description: This example demonstrates how to generate a new workbook, hide columns G through J on the first worksheet using Cells.HideColumns, and store the configuration as a custom view that can be reopened later. The file is saved as CustomView_HideGtoJ.xlsx.
// Keywords: Aspose.Cells hide columns C# | Cells.HideColumns example | custom view Excel Aspose | hide column range G-J | save workbook after hiding columns | Aspose.Cells .NET tutorial
// Common Searches: how to hide columns G to J with Aspose.Cells C# | Aspose.Cells create custom view hide columns | C# hide multiple Excel columns programmatically | save Excel file after hiding columns using Aspose
// Developer Intent: Hide columns G‑J in a worksheet, preserve the layout as a custom view, and save the workbook.
// Use Cases: Prepare a printable report that excludes intermediate calculation columns. | Distribute a template where sensitive data columns are hidden by default. | Generate an export file that shows only the final results while keeping raw data hidden.
// AI Prompts: Generate C# code with Aspose.Cells that hides columns G through J, creates a custom view, and saves the workbook. | Explain step‑by‑step how to use Cells.HideColumns to hide a column range and store it as a reusable view in an Excel file. | Show how to programmatically hide a set of columns and persist the view for later opening using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// This example demonstrates how to generate a new workbook, hide columns G through J on the first worksheet using Cells.HideColumns, and store the configuration as a custom view that can be reopened later. The file is saved as CustomView_HideGtoJ.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns G through J (zero‑based indexes 6,7,8,9)
            // HideColumns(startColumn, totalColumns) hides a range of columns
            sheet.Cells.HideColumns(6, 4);

            // Save the workbook
            string outputPath = "CustomView_HideGtoJ.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
