using Rhino.Geometry;
using Rhino.NodeInCode;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace TJADSZY.genSchool
{
    internal class utilities
    {
        static double outlineGap = 4;
        static double teachingBDWidth = 13;
        static double otherBDgap = 10;
        static double BDgap = 10;
        static double bigBDgap = 30;
        static double otherMinLen = 30;
        static double breakLenOther = 100;

        public static Curve polyline(List<Point3d> inputPts, bool isClosed)
        {
            var tmpFunc = Components.FindComponent("PolyLine").Delegate as dynamic;
            return tmpFunc(inputPts, isClosed)[0];
        }

        public static double area(Curve inputBound)
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

        public static int studentNum(List<Point3d> nakedPts, double PR, Point3d cenPGPt)
        {
            int tmp = 0;

            double totalArea = (area(polyline(nakedPts, true)) - cenPGPt.X * cenPGPt.Y * 4) * PR;
            tmp = Convert.ToInt32(totalArea / (7 + 0.7 + 1.2 + 1.2));
            return tmp;
        }

        public static List<Curve> PGreturn(List<Curve> PG, List<Point3d> pts)
        {

            var tmpFunc = Components.FindComponent("Move").Delegate as dynamic;

            Vector3d trans = new Vector3d();

            trans = new Vector3d(pts[2].X + outlineGap, pts[2].Y + outlineGap, 0);
            var tmp = tmpFunc(PG, trans)[0] as IList<object>;
            return tmp.OfType<Curve>().ToList();
        }

        public static ArrayList PGrelated(List<Curve> s, List<Curve> m, List<Curve> l, List<Point3d> pts)
        {

            Point3d basePtS = cenPt(s[0]);
            Point3d basePtM = cenPt(m[0]);
            Point3d basePtL = cenPt(l[0]);

            ArrayList tmp = new ArrayList();

            if (pts[1].Y - pts[2].Y <= 142.55 + 8)
            {
                tmp.Add(basePtS);
                tmp.Add(s);
            }
            else if (pts[1].Y - pts[2].Y > 142.55 + 8 && pts[1].Y - pts[2].Y <= 178.03 + 8)

            {
                tmp.Add(basePtM);
                tmp.Add(m);
            }
            else
            {
                tmp.Add(basePtL);
                tmp.Add(l);
            }
            return tmp;
        }

        public static List<Box> TeachBuilding(int studentNum, double teachingWidth, List<Point3d> pts, double gapBetweenTeaching, int numBuilding)
        {
            List<Box> breps = new List<Box>();

            breps.Add(new Box(new BoundingBox(new Point3d(pts[3].X - outlineGap - teachingWidth - teachingBDWidth, pts[3].Y+outlineGap, pts[3].Z), new Point3d(pts[3].X - outlineGap - teachingWidth, pts[3].Y + outlineGap + numBuilding * teachingBDWidth + (numBuilding - 1) * gapBetweenTeaching, pts[3].Z+4*4))));
            for (int i = 0; i < numBuilding; i++)
            {
                breps.Add(new Box(new BoundingBox(new Point3d(pts[3].X - outlineGap - teachingWidth, pts[3].Y + outlineGap + i * (gapBetweenTeaching + teachingBDWidth), pts[3].Z), new Point3d(pts[3].X - outlineGap, pts[3].Y + outlineGap + i * (gapBetweenTeaching + teachingBDWidth) + teachingBDWidth, pts[3].Z + 4 * 4))));
            }

            return breps;
        }

        public static List<Box> OtherBuildings(int studentNum, double teachingWidthAll, List<Point3d> pts, Point3d cenPGPt, double teachingLen)
        {
            List<Box> breps = new List<Box>();
            double otherArea = studentNum * 7;

            double leftLenPG = pts[1].Y - pts[2].Y - outlineGap - outlineGap - BDgap - cenPGPt.Y * 2;
            double leftLenTeaching = pts[1].Y - pts[2].Y - outlineGap - outlineGap - BDgap - teachingLen;
            double midWidthRemain = pts[0].X - pts[1].X - outlineGap * 2 - cenPGPt.X * 2 - teachingWidthAll - teachingBDWidth - BDgap * 2;


            if (leftLenPG >= otherMinLen || leftLenTeaching >= otherMinLen)
            {
                int floorNumNorm = Math.Min(Convert.ToInt32(otherArea / ((pts[0].X - pts[1].X - outlineGap * 2) * otherMinLen)), 3);

                if (leftLenPG>=otherMinLen)
                {
                    double otherWidPG = cenPGPt.X * 2;
                    int bdNumOtherPG = Convert.ToInt32(Math.Ceiling(otherWidPG / breakLenOther));
                    for (int i = 0; i < bdNumOtherPG; i++)
                    {
                        breps.Add(new Box(new BoundingBox(new Point3d(pts[2].X + outlineGap + i * (otherWidPG / bdNumOtherPG - BDgap) + BDgap * (i + 0.5), pts[1].Y - outlineGap - leftLenPG, pts[3].Z), new Point3d(pts[2].X + outlineGap + (i + 1) * (otherWidPG / bdNumOtherPG - BDgap) + BDgap * (i + 0.5), pts[1].Y - outlineGap, pts[3].Z + floorNumNorm * 4))));
                    }
                }

                if (leftLenTeaching >= otherMinLen)
                {
                    double otherWidTeaching = teachingWidthAll + teachingBDWidth;
                    int bdNumOtherTch = Convert.ToInt32(Math.Ceiling(otherWidTeaching / breakLenOther));
                    for (int i = 0; i < bdNumOtherTch; i++)
                    {
                        breps.Add(new Box(new BoundingBox(new Point3d(pts[0].X - outlineGap - otherWidTeaching + i * (otherWidTeaching / bdNumOtherTch - BDgap) + BDgap * (i + 0.5), pts[1].Y - outlineGap - leftLenTeaching, pts[3].Z), new Point3d(pts[0].X - outlineGap - otherWidTeaching + (i + 1) * (otherWidTeaching / bdNumOtherTch - BDgap) + BDgap * (i + 0.5), pts[1].Y - outlineGap, pts[3].Z + floorNumNorm * 4))));
                    }
                }

                if (midWidthRemain >= otherMinLen)
                {
                    int bdNumOtherRem = Convert.ToInt32(Math.Ceiling(midWidthRemain / breakLenOther));
                    for (int i = 0; i < bdNumOtherRem; i++)
                    {
                        breps.Add(new Box(new BoundingBox(new Point3d(pts[2].X + outlineGap + cenPGPt.X * 2 + BDgap + i * (midWidthRemain / bdNumOtherRem - BDgap) + BDgap * (i + 0.5), pts[1].Y - outlineGap - Math.Max(leftLenPG, leftLenTeaching), pts[3].Z), new Point3d(pts[2].X + outlineGap + cenPGPt.X * 2 + BDgap + (i + 1) * (midWidthRemain / bdNumOtherRem - BDgap) + BDgap * (i + 0.5), pts[1].Y - outlineGap, pts[3].Z + floorNumNorm * 4))));
                    }
                }
                return breps;
            }

            else if (midWidthRemain >= 20)
            {
                double midLenAll = pts[1].Y - pts[2].Y - outlineGap * 2;

                int floorNumMid = Math.Min(Convert.ToInt32(otherArea / (midLenAll - bigBDgap * 2)), 2);

                breps.Add(new Box(new BoundingBox(new Point3d(pts[2].X + outlineGap + cenPGPt.X * 2 + BDgap, pts[2].Y + outlineGap, pts[3].Z), new Point3d(pts[3].X - outlineGap - teachingWidthAll - teachingBDWidth - BDgap, pts[2].Y + outlineGap + (midLenAll - bigBDgap * 2) / 4, pts[3].Z + floorNumMid * 4))));
                breps.Add(new Box(new BoundingBox(new Point3d(pts[2].X + outlineGap + cenPGPt.X * 2 + BDgap, pts[2].Y + outlineGap + bigBDgap + (midLenAll - bigBDgap * 2) / 4, pts[3].Z), new Point3d(pts[3].X - outlineGap - teachingWidthAll - teachingBDWidth - BDgap, pts[2].Y + outlineGap + bigBDgap + (midLenAll - bigBDgap * 2) / 4 * 3, pts[3].Z + floorNumMid * 4))));
                breps.Add(new Box(new BoundingBox(new Point3d(pts[2].X + outlineGap + cenPGPt.X * 2 + BDgap, pts[2].Y + outlineGap + bigBDgap*2 + (midLenAll - bigBDgap * 2) / 4 *3, pts[3].Z), new Point3d(pts[3].X - outlineGap - teachingWidthAll - teachingBDWidth - BDgap, pts[2].Y + outlineGap + bigBDgap * 2 + (midLenAll - bigBDgap * 2), pts[3].Z + floorNumMid * 4))));

                return breps;
            }
            else
            {
                return null;
            }

        }

        public static ArrayList allBuildings(List<Curve> s, List<Curve> m, List<Curve> l, List<Point3d> nakedPts, double PR, double dayFactor)
        {
            ArrayList buildings = new ArrayList();
            #region set student numbers & total area & playgroundCenPoint
            List<Curve> PGcrv = new List<Curve>();
            Point3d cenPtPG = new Point3d();
            if (PGrelated(s, m, l, nakedPts)[0] is Point3d)
            {
                cenPtPG = (Point3d)PGrelated(s, m, l, nakedPts)[0];
            }
            if (PGrelated(s, m, l, nakedPts)[1] is List<Curve>)
            {
                PGcrv = (List<Curve>)PGrelated(s, m, l, nakedPts)[1];
            }
            double totalArea = (area(polyline(nakedPts, true)) - cenPtPG.X * cenPtPG.Y * 4) * PR;
            int SNum = Convert.ToInt32(totalArea / (7 + 0.7 + 1.2 + 1.2));
            double teachingWidthAll = Math.Min(nakedPts[3].X - nakedPts[2].X - outlineGap - outlineGap - teachingBDWidth - outlineGap - cenPtPG.X * 2, 150);
            double teachingArea = SNum * 7;
            double gapBetweenTeaching = Math.Max(dayFactor * 4 * 4, 25);
            int teachingNumBuilding = Convert.ToInt32(teachingArea / (teachingWidthAll * teachingBDWidth * 4));
            double teachingLen = teachingBDWidth * teachingNumBuilding + gapBetweenTeaching * (teachingNumBuilding - 1);

            #endregion

            buildings.Add(PGreturn(PGcrv, nakedPts));
            buildings.Add(TeachBuilding(SNum, teachingWidthAll, nakedPts, gapBetweenTeaching, teachingNumBuilding));
            buildings.Add(OtherBuildings(SNum, teachingWidthAll, nakedPts, cenPtPG, teachingLen));

            return buildings;
        }

    }
}
