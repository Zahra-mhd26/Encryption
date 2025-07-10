using Grasshopper.Kernel;
using System;
using System.IO; // Required for file operations
using System.Security.Cryptography; // Required for encryption

// TODO: If ClosedXML is used for specific Excel manipulation before encryption, add:
// using ClosedXML.Excel;

namespace ExcelEncryption
{
    public class ExcelFileEncryptorComponent : GH_Component
    {
        public ExcelFileEncryptorComponent()
          : base("Excel File Encryptor", "ExcelEncrypt",
                "Encrypts an Excel file using AES encryption.",
                "Utilities", "Encryption") // Matched category/subcategory of old component
        {
        }

        public override Guid ComponentGuid => new Guid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF"); // New GUID

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Input Excel File Path", "I", "The path of the Excel file to encrypt", GH_ParamAccess.item);
            pManager.AddTextParameter("Output Encrypted File Path", "O", "The path where the encrypted file will be saved", GH_ParamAccess.item);
            // Future enhancement: Add a parameter to trigger a Windows Form for file selection or options
            // pManager.AddBooleanParameter("Show Options", "Opt", "Show encryption options form", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddBooleanParameter("Success", "S", "True if the file was successfully encrypted.", GH_ParamAccess.item);
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
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Input file path is invalid or file does not exist.");
                DA.SetData(0, false);
                return;
            }

            // Validate output path
            if (string.IsNullOrEmpty(outputFilePath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Output file path is invalid.");
                DA.SetData(0, false);
                return;
            }

            // It's good practice to ensure the output directory exists.
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

            // Placeholder for Windows Form interaction if needed:
            // bool showOptions = false;
            // DA.GetData("Show Options", ref showOptions); // Assuming "Show Options" is the 3rd input if enabled
            // if (showOptions)
            // {
            //     // Launch form here, potentially modifying inputFilePath or outputFilePath
            //     // For example:
            //     // var form = new EncryptionOptionsForm(inputFilePath, outputFilePath);
            //     // if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //     // {
            //     //     inputFilePath = form.SelectedInputPath;
            //     //     outputFilePath = form.SelectedOutputPath;
            //     // }
            //     // else return; // User cancelled
            // }

            try
            {
                EncryptFile(inputFilePath, outputFilePath);
                success = true;
            }
            catch (Exception ex)
            {
                success = false;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Encryption failed: {ex.Message}");
            }

            DA.SetData(0, success);
        }


    // --- Start of Encryption Logic (copied from Encryption_old) ---

    // IMPORTANT: These Key and IV are hardcoded and should be identical to the ones
    // in your DecryptionComponent if you want to decrypt these files.
    // For production systems, consider more secure ways to manage keys.
    private static readonly byte[] Key = Convert.FromBase64String("GZuHOfU1C5F6SgI+PRKo+g=="); // 16 bytes for AES-128. Ensure this matches DecryptionComponent.
    private static readonly byte[] IV = Convert.FromBase64String("0v2u51jxM1dwalKzbWiZHg=="); // 16 bytes for AES. Ensure this matches DecryptionComponent.

    private void EncryptFile(string inputFile, string outputFile)
    {
        // The old project used Aes.Create() which defaults to AES-256 if key is 32 bytes,
        // but the key here is 16 bytes, implying AES-128.
        // Let's be explicit with AesCryptoServiceProvider for clarity if needed,
        // or ensure Aes.Create() is configured correctly if issues arise.
        // For now, direct copy of Aes.Create() is fine as it should work.
        using (Aes aes = Aes.Create())
        {
            aes.Key = Key;
            aes.IV = IV;

            using (FileStream fsInput = new FileStream(inputFile, FileMode.Open, FileAccess.Read))
            using (FileStream fsEncrypted = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
            using (ICryptoTransform encryptor = aes.CreateEncryptor(aes.Key, aes.IV))
            using (CryptoStream csEncrypt = new CryptoStream(fsEncrypted, encryptor, CryptoStreamMode.Write))
            {
                fsInput.CopyTo(csEncrypt);
            }
        }
    }

    // Optional: If you need to generate new keys/IVs (e.g., for a key management system)
    // You can include these helper methods. For this task, we use the hardcoded ones.
    /*
    private static byte[] GenerateRandomKey()
    {
        using (var rng = new RNGCryptoServiceProvider())
        {
            byte[] key = new byte[16]; // 128-bit key (or 32 for AES-256)
            rng.GetBytes(key);
            return key;
        }
    }

    private static byte[] GenerateRandomIV()
    {
        using (var rng = new RNGCryptoServiceProvider())
        {
            byte[] iv = new byte[16]; // 128-bit IV
            rng.GetBytes(iv);
            return iv;
        }
    }
    */
    // --- End of Encryption Logic ---

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => null; // Or add a custom icon
    }
}