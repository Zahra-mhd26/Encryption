using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace ExcelEncryption
{
    public class ExcelEncryptionInfo : GH_AssemblyInfo
    {
        public override string Name => "ExcelEncryption";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "";

        public override Guid Id => new Guid("24f91c2c-41d2-4634-882a-9d000bcb9180");

        //Return a string identifying you or your company.
        public override string AuthorName => "";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "";
    }
}