// Title: Aspose.Cells .NET: Load Workbook, Enable ISO 29500 Strict Mode, and Handle Save Permission Errors
// Description: Loads an existing Excel file with Aspose.Cells, sets OoxmlCompliance.Iso29500_2008_Strict, and attempts to save to a protected directory (e.g., C:\Windows\System32). The sample catches the resulting UnauthorizedAccessException and displays the error message.
// Keywords: Aspose.Cells | C# | strict compliance | ISO 29500 | OoxmlCompliance | Workbook.Save | permission error | UnauthorizedAccessException | protected folder | Windows System32
// Common Searches: Aspose.Cells set strict OOXML compliance C# | How to catch permission error when saving Excel with Aspose.Cells | Saving workbook to C:\Windows\System32 using Aspose.Cells | Exception thrown for insufficient write permissions Aspose.Cells | Enable ISO 29500 strict mode before saving Aspose.Cells
// Developer Intent: The developer wants to confirm that enabling ISO/IEC 29500:2008 strict compliance does not override file‑system security and to implement robust error handling for save operations targeting directories without write access.
// Use Cases: Validate that strict OOXML compliance is applied before persisting a workbook. | Test application behavior when attempting to save to a location that requires administrative rights. | Capture and log UnauthorizedAccessException for user feedback or audit purposes. | Demonstrate that compliance settings do not suppress permission‑related exceptions.
// AI Prompts: Generate C# Aspose.Cells code that loads a workbook, sets OoxmlCompliance.Iso29500_2008_Strict, and saves to a folder with limited permissions while handling UnauthorizedAccessException. | Explain how Aspose.Cells strict compliance mode interacts with the .NET file‑system security model during a save operation. | What are best practices for handling save failures caused by insufficient permissions in Aspose.Cells applications?

using System;
using Aspose.Cells;

// Loads an existing Excel file with Aspose.Cells, sets OoxmlCompliance.Iso29500_2008_Strict, and attempts to save to a protected directory (e.g., C:\Windows\System32). The sample catches the resulting UnauthorizedAccessException and displays the error message.
class StrictComplianceSaveDemo
{
    static void Main()
    {
        // Load an existing workbook (ensure the file exists at this path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Enable ISO/IEC 29500:2008 Strict compliance for OOXML
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Define a path that typically requires elevated permissions
        string restrictedPath = @"C:\Windows\System32\restricted.xlsx";

        try
        {
            // Attempt to save the workbook to the restricted location
            workbook.Save(restrictedPath);
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            // Expected failure due to insufficient permissions
            Console.WriteLine("Failed to save workbook: " + ex.Message);
        }
    }
}
