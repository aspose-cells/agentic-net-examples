// Title: Check if an Excel workbook’s structure or window is protected using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file with Aspose.Cells, reads the WorkbookSettings.IsProtected flag to determine whether the workbook’s structure or window is protected, and writes the boolean result to the console.
// Keywords: Aspose.Cells | C# | WorkbookSettings.IsProtected | check workbook protection | Excel structure protection | window protection | detect workbook protection | read protection flag | Aspose.Cells example
// Common Searches: Aspose.Cells how to check workbook protection | C# get workbook structure protection status | IsProtected property Aspose.Cells | determine if Excel workbook is locked with Aspose.Cells | read workbook protection flag C#
// Developer Intent: Identify whether the loaded workbook’s structure or window is protected.
// Use Cases: Prevent editing operations when a workbook is protected. | Show or hide UI controls (Save, Modify) based on protection state. | Log protection status of incoming workbooks for compliance audits.
// AI Prompts: Generate C# code that uses Aspose.Cells to unprotect a workbook when WorkbookSettings.IsProtected is true, handling a password. | Show how to differentiate between structure protection and window protection using Aspose.Cells properties. | Explain how to programmatically remove workbook protection with a password in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace WorkbookProtectionCheck
{
    // Loads an Excel file with Aspose.Cells, reads the WorkbookSettings.IsProtected flag to determine whether the workbook’s structure or window is protected, and writes the boolean result to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook to be examined
            string filePath = "SampleWorkbook.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(filePath);

            // Access workbook settings
            WorkbookSettings settings = workbook.Settings;

            // Determine if the workbook structure or window is protected
            bool isProtected = settings.IsProtected;

            // Output the protection status to the console
            Console.WriteLine("Workbook structure/window protected: " + isProtected);
        }
    }
}
