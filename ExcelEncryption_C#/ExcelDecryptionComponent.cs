using Grasshopper.Kernel;
using System;
using System.IO; // Required for file operations
using System.Security.Cryptography; // Required for decryption

namespace ExcelEncryption
{
    public class ExcelFileDecryptorComponent : GH_Component
    {
        public ExcelFileDecryptorComponent()
          : base("Excel File Decryptor", "ExcelDecrypt",
                "Decrypts an Excel file using AES encryption.",
                "Utilities", "Encryption") // Matched category/subcategory
        {
        }

        // Create a new unique GUID for this component
        public override Guid ComponentGuid => new Guid("B2C3D4E5-F6A7-8901-2345-678901BCDEFA");

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Input Encrypted File Path", "I", "The path of the encrypted file to decrypt", GH_ParamAccess.item);
            pManager.AddTextParameter("Output Decrypted File Path", "O", "The path where the decrypted file will be saved", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Success", "S", "True if the file was successfully decrypted.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string inputFilePath = string.Empty;
            string outputFilePath = string.Empty;
            bool success = false;

            if (!DA.GetData(0, ref inputFilePath)) return;
            if (!DA.GetData(1, ref outputFilePath)) return;

            // Validate input path
            if (string.IsNullOrEmpty(inputFilePath) || !File.Exists(inputFilePath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Input encrypted file path is invalid or file does not exist.");
                DA.SetData(0, false);
                return;
            }

            // Validate output path
            if (string.IsNullOrEmpty(outputFilePath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Output decrypted file path is invalid.");
                DA.SetData(0, false);
                return;
            }

            // Ensure the output directory exists
            try
            {
                string outputDirectory = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error creating output directory: {ex.Message}");
                DA.SetData(0, false);
                return;
            }

            try
            {
                DecryptFile(inputFilePath, outputFilePath);
                success = true;
            }
            catch (CryptographicException cex) // Catch specific crypto errors for better feedback
            {
                success = false;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Decryption failed. This could be due to an incorrect key/IV, corrupted file, or if the file was not encrypted with the corresponding encryptor: {cex.Message}");
            }
            catch (Exception ex)
            {
                success = false;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Decryption failed: {ex.Message}");
            }

            DA.SetData(0, success);
        }

        // --- Start of Decryption Logic ---

        // IMPORTANT: These Key and IV MUST be identical to the ones
        // in your EncryptionComponent.
        private static readonly byte[] Key = Convert.FromBase64String("GZuHOfU1C5F6SgI+PRKo+g=="); // Same as EncryptionComponent
        private static readonly byte[] IV = Convert.FromBase64String("0v2u51jxM1dwalKzbWiZHg=="); // Same as EncryptionComponent

        private void DecryptFile(string inputFile, string outputFile)
        {
            using (Aes aes = Aes.Create())
            {
                aes.Key = Key;
                aes.IV = IV;

                using (FileStream fsEncrypted = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
                using (FileStream fsDecrypted = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                using (ICryptoTransform decryptor = aes.CreateDecryptor(aes.Key, aes.IV))
                using (CryptoStream csDecrypt = new CryptoStream(fsEncrypted, decryptor, CryptoStreamMode.Read))
                {
                    csDecrypt.CopyTo(fsDecrypted);
                }
            }
        }
        // --- End of Decryption Logic ---

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => null; // Or add a custom icon
    }
}
