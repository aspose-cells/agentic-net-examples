using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsExternalConnectionValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add a dummy external connection for demonstration purposes
            // In real scenarios the workbook would already contain connections
            // ExternalConnection connection = workbook.DataConnections.Add("MyConnection", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=sample.xlsx;", true);
            // connection.Credentials = CredentialsMethodType.Integrated; // set credentials as needed

            // Validate that each external connection has non‑empty credentials before saving
            ValidateExternalConnections(workbook);

            // Save the workbook after successful validation
            workbook.Save("ValidatedWorkbook.xlsx");
        }

        /// <summary>
        /// Checks every external connection in the workbook to ensure credentials are specified.
        /// Throws an exception if any connection lacks credentials.
        /// </summary>
        /// <param name="workbook">The workbook to validate.</param>
        private static void ValidateExternalConnections(Workbook workbook)
        {
            // Iterate through all external connections
            for (int i = 0; i < workbook.DataConnections.Count; i++)
            {
                ExternalConnection conn = workbook.DataConnections[i];

                // The Credentials property is an enum; treat 'None' as empty credentials.
                // If the property is obsolete, it still reflects the authentication method.
                if (conn.Credentials == CredentialsMethodType.None)
                {
                    throw new InvalidOperationException(
                        $"External connection '{conn.Name}' does not have credentials set.");
                }

                // Additionally, ensure that the password is saved if required.
                // If the connection uses a password, SavePassword should be true.
                // This check is optional and can be adjusted based on specific needs.
                // if (conn.SavePassword == false && /* condition indicating password is needed */)
                // {
                //     throw new InvalidOperationException(
                //         $"External connection '{conn.Name}' requires the password to be saved.");
                // }
            }
        }
    }
}