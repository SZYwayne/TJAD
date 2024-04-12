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

        public static Point3d cenPt(Curve inputCrv)
        {
            var tmpFunc = Components.FindComponent("Area").Delegate as dynamic;
            var tmpList = tmpFunc(inputCrv)[1] as IList<object>;
            List<Point3d> outputList = tmpList.OfType<Point3d>().ToList();

            return outputList[0];
        }

        public static List<Curve> PGreturn(int studentNum, List<Curve> s, List<Curve> m, List<Curve> l, List<Point3d> pts)
        {
            Point3d basePtS = cenPt(s[0]);
            Point3d basePtM = cenPt(m[0]);
            Point3d basePtL = cenPt(l[0]);

            if (studentNum < 45 * 24) return s;
            else if (studentNum < 45 * 36 & studentNum >= 24 * 45) return m; 
            else return l;
        }

        public static Point3d test(int studentNum, List<Curve> s, List<Curve> m, List<Curve> l, List<Point3d> pts)
        {
            Point3d basePtS = cenPt(s[0]);
            Point3d basePtM = cenPt(m[0]);
            Point3d basePtL = cenPt(l[0]);

            if (studentNum < 45 * 24) return basePtS;
            else if (studentNum < 45 * 36 & studentNum >= 24 * 45) return basePtM;
            else return basePtL;
        }
    }
}
