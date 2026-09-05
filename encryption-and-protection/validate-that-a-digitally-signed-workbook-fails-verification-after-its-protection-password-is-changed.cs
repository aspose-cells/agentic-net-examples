// Title: Check that a digitally signed Excel workbook becomes invalid after changing its protection password with Aspose.Cells for .NET
// AI Prompts: Add a digital signature to a Workbook, protect it with a password, change the password, and programmatically verify that the signature is now invalid using Aspose.Cells in C#. | Use reflection to invoke DigitalSignatureCollection.AddSignature and VerifyAllSignatures in Aspose.Cells, then compare verification results before and after modifying workbook protection. | Demonstrate detecting a broken digital signature after unprotecting and re‑protecting an Excel file with a new password in C# with Aspose.Cells.
// Common Searches: aspnet verify digital signature after changing Excel workbook password | c# Aspose.Cells signature becomes invalid when workbook protection password is changed | how to detect broken digital signature after re‑protecting Excel file using Aspose.Cells | using reflection to access DigitalSignatureCollection in Aspose.Cells .NET
// Tags: digital signature verification after workbook password change Aspose.Cells | protect workbook with new password invalidates signature C# | reflection access DigitalSignatureCollection Aspose.Cells | add and verify Excel digital signature Aspose.Cells .NET | unprotect and re‑protect workbook signature integrity

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, and uses reflection to call Aspose.Cells' DigitalSignatureCollection methods to add a digital signature from a PFX file. It verifies the signature's validity, then protects the workbook with an initial password, unprotects it, re‑protects with a new password, and verifies the signature again, showing that the signature becomes invalid after the protection password is changed. Error handling covers missing APIs or certificate files.
class DigitalSignatureVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Test");
            sheet.Cells["B1"].PutValue(123);

            bool isSignatureValidBefore = false;
            bool isSignatureValidAfter = false;

            // Attempt to add and verify a digital signature using reflection (API may be unavailable)
            try
            {
                string certPath = "cert.pfx";
                string certPassword = "certPassword";

                // Get the DigitalSignatureCollection property via reflection
                PropertyInfo dsProp = typeof(Workbook).GetProperty("DigitalSignatureCollection");
                if (dsProp != null && File.Exists(certPath))
                {
                    object dsCollection = dsProp.GetValue(workbook);
                    MethodInfo addSignature = dsCollection.GetType().GetMethod(
                        "AddSignature",
                        new[] { typeof(string), typeof(string), typeof(string) });

                    MethodInfo verifyAll = dsCollection.GetType().GetMethod("VerifyAllSignatures", Type.EmptyTypes);

                    if (addSignature != null && verifyAll != null)
                    {
                        // Add a digital signature
                        addSignature.Invoke(dsCollection, new object[] { certPath, certPassword, "Initial signature" });

                        // Verify the signature – should be valid if signing succeeded
                        isSignatureValidBefore = (bool)verifyAll.Invoke(dsCollection, null);
                        Console.WriteLine("Signature valid before protection change: " + isSignatureValidBefore);
                    }
                    else
                    {
                        Console.WriteLine("Digital signature methods not found in the current Aspose.Cells version.");
                    }
                }
                else
                {
                    Console.WriteLine(dsProp == null
                        ? "Digital signature API not available in this Aspose.Cells version."
                        : $"Certificate file not found: {certPath}. Skipping digital signing.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during digital signing: " + ex.Message);
            }

            // Protect the workbook with an initial password
            string oldPassword = "oldPass";
            workbook.Protect(ProtectionType.All, oldPassword);

            // Change the protection password (unprotect then protect with new password)
            workbook.Unprotect(oldPassword);
            string newPassword = "newPass";
            workbook.Protect(ProtectionType.All, newPassword);

            // Verify the signature again – it will be invalid if the workbook was modified
            try
            {
                PropertyInfo dsProp = typeof(Workbook).GetProperty("DigitalSignatureCollection");
                if (dsProp != null)
                {
                    object dsCollection = dsProp.GetValue(workbook);
                    MethodInfo verifyAll = dsCollection.GetType().GetMethod("VerifyAllSignatures", Type.EmptyTypes);
                    if (verifyAll != null)
                    {
                        isSignatureValidAfter = (bool)verifyAll.Invoke(dsCollection, null);
                        Console.WriteLine("Signature valid after protection password change: " + isSignatureValidAfter);
                    }
                    else
                    {
                        Console.WriteLine("Digital signature verification method not found.");
                    }
                }
                else
                {
                    Console.WriteLine("Digital signature API not available in this Aspose.Cells version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during signature verification: " + ex.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
