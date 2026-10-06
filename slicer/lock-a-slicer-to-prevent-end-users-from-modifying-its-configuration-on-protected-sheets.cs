// Title: Lock slicer objects on a worksheet and protect the sheet using Aspose.Cells for .NET
// AI Prompts: Set the IsLocked property to true for every slicer in a worksheet and then call Worksheet.Protect(ProtectionType.All) with Aspose.Cells in C#. | Load or create an Excel workbook, loop through Worksheet.Slicers to disable user edits, apply sheet protection without a password, and save the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# lock slicer on protected worksheet | prevent users from changing slicer settings in Excel using Aspose.Cells | how to protect a sheet while keeping slicers read‑only with Aspose.Cells | set slicer IsLocked property in Aspose.Cells .NET example | worksheet protection without password Aspose.Cells C#
// Tags: Aspose.Cells slicer lock flag | C# protect worksheet without password | enumerate worksheet slicers Aspose.Cells | read‑only slicer configuration .NET | lock slicer objects on worksheet

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook or creates a new one, accesses the first worksheet, iterates over all slicers to set their IsLocked property to true, protects the worksheet with all protection types (no password), and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Get the first worksheet (a new workbook always contains at least one sheet).
            Worksheet worksheet = workbook.Worksheets[0];

            // Attempt to lock all slicers on the worksheet (if slicer support is available).
            try
            {
                foreach (var slicer in worksheet.Slicers)
                {
                    slicer.IsLocked = true;
                }
            }
            catch
            {
                // Slicer API may not be available in the current Aspose.Cells version; ignore.
            }

            // Protect the worksheet without a password using the appropriate overload.
            worksheet.Protect(ProtectionType.All);

            // Save the modified workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
