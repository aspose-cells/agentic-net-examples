// Title: How to lock an Excel table (ListObject) to prevent users from adding rows with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to protect a worksheet and disable row insertion for a specific ListObject. | Show how to apply worksheet protection that blocks adding rows to an Excel table while still allowing other edits. | Generate an example that saves a workbook after preventing expansion of a ListObject in C#. | Provide a snippet that sets a password‑protected worksheet protection to enforce a fixed table size with Aspose.Cells.
// Common Searches: Aspose.Cells C# prevent users from adding rows to an Excel table | How to disable row insertion in a ListObject with Aspose.Cells .NET | Lock Excel table size using worksheet protection in Aspose.Cells | C# code to stop expanding a ListObject after loading a workbook with Aspose.Cells | Set worksheet protection to block table row addition in Aspose.Cells for .NET
// Tags: Aspose.Cells disable ListObject row addition | C# enforce fixed Excel table size | apply sheet lock to stop table growth | Aspose.Cells lock table expansion | prevent table row insertion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// The example loads an existing workbook, locates the first ListObject (Excel table), applies full worksheet protection to block row insertion (including table expansion), optionally sets a password, and saves the protected workbook to a new file.
class DisableListObjectRowAddition
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one ListObject (Excel Table)
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("Warning: No tables (ListObjects) were found on the first worksheet.");
            }
            else
            {
                // Retrieve the first ListObject (Excel Table) on the worksheet
                ListObject table = sheet.ListObjects[0];
                // Additional table-specific logic could be placed here if needed
            }

            // Protect the worksheet to prevent row insertion and other modifications
            sheet.Protect(ProtectionType.All);
            // Optionally, set a password for the protection
            // sheet.Protection.SetPassword("yourPassword");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
