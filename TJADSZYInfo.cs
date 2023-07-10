using Grasshopper;
using Grasshopper.Kernel;
using System;
using System.Drawing;

namespace TJADSZY
{
    public class TJADSZYInfo : GH_AssemblyInfo
    {
        public override string Name => "TJADSZY";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "";

        public override Guid Id => new Guid("74367fb9-a20d-40bb-8ce9-35e80d9b751f");

        //Return a string identifying you or your company.
        public override string AuthorName => "";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "";
    }
}