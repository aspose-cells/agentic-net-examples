// Title: Load an Excel workbook from a UNC network share, detect encryption, and handle access‑permission errors with Aspose.Cells for .NET
// Common Searches: asp.net load excel file from UNC path and check if it is password protected with Aspose.Cells | catch CellsException to determine encrypted workbook when opening file on network share | handle UnauthorizedAccessException while opening Excel workbook from shared folder using Aspose.Cells | verify existence of Excel file on network share before loading with Aspose.Cells C# | determine encryption status of workbook without password using Aspose.Cells
// Tags: remote share workbook loading Aspose.Cells | encryption status check Aspose.Cells | shared folder permission handling Aspose.Cells | CellsException handling for encrypted files | pre‑load file presence check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks that an Excel file exists on a UNC network share, attempts to open it with Aspose.Cells, treats any CellsException as an indication that the workbook is encrypted, catches UnauthorizedAccessException to report permission problems, and finally prints whether the workbook is encrypted.
class Program
{
    static void Main()
    {
        // UNC path to the workbook on a network share
        string networkPath = @"\\ServerName\ShareFolder\Workbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(networkPath))
        {
            Console.WriteLine($"File not found: {networkPath}");
            return;
        }

        bool isEncrypted = false;
        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(networkPath);
        }
        catch (CellsException ce)
        {
            // If the file is encrypted, Aspose.Cells throws a CellsException.
            // We treat any CellsException here as an indication of encryption.
            isEncrypted = true;
        }
        catch (UnauthorizedAccessException uaEx)
        {
            // Permission issues when accessing the network share
            Console.WriteLine($"Access permission issue: {uaEx.Message}");
            return;
        }
        catch (Exception ex)
        {
            // Other errors that occur during loading
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Output encryption status
        Console.WriteLine($"Workbook encryption status: {(isEncrypted ? "Encrypted" : "Not encrypted")}");
    }
}
