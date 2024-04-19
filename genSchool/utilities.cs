using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rhino.NodeInCode;
using Grasshopper.Kernel.Geometry;
using System.Collections;

namespace TJADSZY.genSchool
{
    internal class utilities
    {
        static double outlineGap = 4;
        static double teachingBDWidth = 13;

        public static Curve polyline(List<Point3d> inputPts, bool isClosed)
        {
            var tmpFunc = Components.FindComponent("PolyLine").Delegate as dynamic;
            return tmpFunc(inputPts, isClosed)[0];
        }

        public static  double area(Curve inputBound)
        {
            var tmpFunc = Components.FindComponent("Area").Delegate as dynamic;
            var tmpList = tmpFunc(inputBound)[0] as IList<object>;
            List<double> outputList = tmpList.OfType<double>().ToList();
            return outputList[0];
        }

        public static Point3d cenPt(Curve inputCrv)
        {
            var tmpFunc = Components.FindComponent("Area").Delegate as dynamic;
            var tmpList = tmpFunc(inputCrv)[1] as IList<object>;
            List<Point3d> outputList = tmpList.OfType<Point3d>().ToList();

            return outputList[0];
        }

        public static int SNum(List<Point3d> nakedPts, double PR, List<Curve> s, List<Curve> m, List<Curve> l)
        {
            int tmp = 0;
            Point3d cenPGPt = PGcenPt(s, m, l, nakedPts);

            double totalArea = (area(polyline(nakedPts, true)) - cenPGPt.X * cenPGPt.Y * 4) * PR;
            tmp = Convert.ToInt32(totalArea / (7 + 0.7 + 1.2 + 1.2));
            return tmp;
        }

        public static List<Curve> PGreturn(List<Curve> s, List<Curve> m, List<Curve> l, List<Point3d> pts)
        {

            var tmpFunc = Components.FindComponent("Move").Delegate as dynamic;

            Point3d basePtS = cenPt(s[0]);
            Point3d basePtM = cenPt(m[0]);
            Point3d basePtL = cenPt(l[0]);

            List<Curve> curves = new List<Curve>();
            Vector3d trans = new Vector3d();

            if (pts[1].Y - pts[2].Y <= 142.55 + 8)
            {
                curves = s;
                trans = new Vector3d(pts[2].X + outlineGap, pts[2].Y + outlineGap, 0);
            }
            else if (pts[1].Y - pts[2].Y > 142.55 + 8 && pts[1].Y - pts[2].Y <= 178.03 + 8)

            {
                curves = m;
                trans = new Vector3d(pts[2].X + outlineGap, pts[2].Y + outlineGap, 0);
            }
            else
            {
                curves = l;
                trans = new Vector3d(pts[2].X + outlineGap, pts[2].Y + outlineGap, 0);
            }
            var tmp = tmpFunc(curves, trans)[0] as IList<object>;
            return tmp.OfType<Curve>().ToList();
        }

        public static Point3d PGcenPt(List<Curve> s, List<Curve> m, List<Curve> l, List<Point3d> pts)
        {

            var tmpFunc = Components.FindComponent("Move").Delegate as dynamic;

            Point3d basePtS = cenPt(s[0]);
            Point3d basePtM = cenPt(m[0]);
            Point3d basePtL = cenPt(l[0]);

            Point3d PGcen = new Point3d();

            if (pts[1].Y - pts[2].Y <= 142.55 + 8)
            {
                PGcen = basePtS;
            }
            else if (pts[1].Y - pts[2].Y > 142.55 + 8 && pts[1].Y - pts[2].Y <= 178.03 + 8)

            {
                PGcen = basePtM;
            }
            else
            {
                PGcen = basePtL;
            }
            return PGcen;
        }

        public static List<Box> TeachBuilding(int studentNum, Point3d centerPoint, List<Point3d> pts, double dayFactor)
        {
            List<Box> breps = new List<Box>();
            double teachingArea = studentNum * 7;

            double teachingWidth = Math.Min(pts[3].X - pts[2].X - outlineGap - outlineGap - teachingBDWidth - outlineGap - centerPoint.X * 2, 100);
            double gapBetween = Math.Max(dayFactor * 4 * 4, 25);
            int numBuilding = Convert.ToInt32(teachingArea / (teachingWidth * teachingBDWidth * 4));

            breps.Add(new Box(new BoundingBox(new Point3d(pts[3].X - outlineGap - teachingWidth - teachingBDWidth, pts[3].Y+outlineGap, pts[3].Z), new Point3d(pts[3].X - outlineGap - teachingWidth, pts[3].Y + outlineGap + numBuilding * teachingBDWidth + (numBuilding - 1) * gapBetween, pts[3].Z+4*4))));
            for (int i = 0; i < numBuilding; i++)
            {
                breps.Add(new Box(new BoundingBox(new Point3d(pts[3].X - outlineGap - teachingWidth, pts[3].Y + outlineGap + i * (gapBetween + teachingBDWidth), pts[3].Z), new Point3d(pts[3].X - outlineGap, pts[3].Y + outlineGap + i * (gapBetween + teachingBDWidth) + teachingBDWidth, pts[3].Z + 4 * 4))));
            }

            return breps;
        }

        public static ArrayList allBuildings()
        {
            ArrayList buildings = new ArrayList();
            #region set student numbers and total area
            int studentNum = 0;



            #endregion




            return buildings;
        }

    }
}
