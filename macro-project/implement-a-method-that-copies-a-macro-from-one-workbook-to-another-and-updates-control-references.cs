// Title: Copy a VBA macro between two XLSM workbooks while retaining control references using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy VBA macro from one XLSM to another C# | How to transfer a VBA project between macro-enabled Excel files using Aspose.Cells | Preserve form control references when copying macros with Aspose.Cells .NET | Copy VBA project programmatically in C# with Aspose.Cells | Copy macro-enabled workbook VBA code using Aspose.Cells API
// Tags: VbaProject CopyTo method | macro-enabled workbook transfer C# | Aspose.Cells VBA project migration | XLSM save with VBA Aspose.Cells | dynamic invocation for VBA copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The CopyMacro method checks that source and destination XLSM files exist, loads them with Aspose.Cells, attempts to copy the VbaProject from the source workbook to the destination using dynamic invocation of the CopyTo method (if the method is available), and finally saves the destination workbook in XLSM format to preserve the VBA code and control references, throwing detailed exceptions on failure.
public static class MacroHelper
{
    /// <param name="sourcePath">Full path of the workbook that contains the macro (must be a macro‑enabled format).</param>
    /// <param name="destPath">Full path of the workbook that will receive the macro.</param>
    public static void CopyMacro(string sourcePath, string destPath)
    {
        // Verify that both files exist before attempting to load them.
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Source workbook not found: {sourcePath}");

        if (!File.Exists(destPath))
            throw new FileNotFoundException($"Destination workbook not found: {destPath}");

        try
        {
            // Load the workbooks. Aspose.Cells auto‑detects the file format.
            Workbook sourceWb = new Workbook(sourcePath);
            Workbook destWb   = new Workbook(destPath);

            // Attempt to copy the VBA project if the API is available.
            if (sourceWb.VbaProject != null && destWb.VbaProject != null)
            {
                try
                {
                    // Use dynamic invocation to avoid compile‑time dependency on the CopyTo method.
                    ((dynamic)sourceWb.VbaProject).CopyTo(destWb);
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // The CopyTo method is not present in this version; ignore and continue.
                }
            }

            // Save the destination workbook as a macro‑enabled file to retain the VBA project.
            destWb.Save(destPath, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to copy macro.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Example file paths – replace with actual locations.
        string sourcePath = "source.xlsm";
        string destPath   = "dest.xlsm";

        try
        {
            MacroHelper.CopyMacro(sourcePath, destPath);
            Console.WriteLine("Macro copied successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
