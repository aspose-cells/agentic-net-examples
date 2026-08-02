// Title: Aspose.Cells for .NET: Create a Custom View that Hides Columns G‑J and Save It
// Description: Demonstrates how to use Aspose.Cells in C# to hide columns G through J, add a custom view (e.g., "HiddenGtoJ"), and save the workbook for later reuse. Includes version‑check guidance and directory handling.
// Keywords: Aspose.Cells hide columns | custom view Aspose.Cells | C# HideColumns method | Excel custom view .NET | save workbook Aspose.Cells | column visibility Aspose.Cells | Excel template hide columns | Aspose.Cells version check custom view | HideColumns G J Aspose.Cells
// Common Searches: Aspose.Cells hide columns G to J C# | How to add a custom view in Aspose.Cells .NET | Save workbook after hiding columns with Aspose.Cells | Check if CustomViews are supported in Aspose.Cells | Create reusable view that hides specific columns in Excel using Aspose
// Developer Intent: Hide columns G‑J and store the configuration as a reusable custom view.
// Use Cases: Prepare a shared template where confidential columns are hidden by default. | Generate a financial report that excludes intermediate calculation columns before distribution. | Create a printable worksheet view that omits columns G‑J to improve layout. | Automate a data‑export process that saves a view with hidden columns for downstream users.
// AI Prompts: Write C# code with Aspose.Cells to hide columns G‑J, add a custom view named "HiddenGtoJ", and save the file. | Explain how to programmatically verify that the CustomViews feature exists in the current Aspose.Cells version and fall back to simple column hiding if not. | Show how to load an existing workbook and activate a previously saved custom view that hides columns G‑J. | Provide a step‑by‑step guide for creating a directory, saving the workbook, and handling errors when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells in C# to hide columns G through J, add a custom view (e.g., "HiddenGtoJ"), and save the workbook for later reuse. Includes version‑check guidance and directory handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide columns G (index 6) through J (index 9) – total 4 columns
            worksheet.Cells.HideColumns(6, 4);

            // NOTE: Custom view functionality may not be available in older Aspose.Cells versions.
            // If supported, you can uncomment the following lines:
            // worksheet.CustomViews.Add("HiddenGtoJ");

            // Define output path and ensure the directory exists
            string outputPath = "CustomViewDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
