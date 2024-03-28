using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rhino.NodeInCode;

namespace TJADSZY.genSchool
{
    internal class utilities
    {
        public static Curve polyline(List<Point3d> inputPts, bool isClosed)
        {
            var tmpFunc = Components.FindComponent("PolyLine").Delegate as dynamic;
            return tmpFunc(inputPts, isClosed)[0];
        }

        public static  List<double> area(Curve inputBound)
        {
            var tmpFunc = Components.FindComponent("Area").Delegate as dynamic;
            var tmpList = tmpFunc(inputBound)[0] as IList<object>;
            List<double> outputList = tmpList.OfType<double>().ToList();
            return outputList;
        }
    }
}
