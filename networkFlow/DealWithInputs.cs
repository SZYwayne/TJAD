using Rhino.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TJADSZY.networkFlow
{
    internal static class DealWithInputs
    {
        public static List<Point3d> noRepeatPts(List<Point3d> pts)
        {
            List<Point3d> tmpPts = pts;

            for (int i = 0; i < pts.Count - 1; i++)
            {
                for (int j = i + 1; i < pts.Count; j++)
                {
                    if (pts[i] == pts[j])
                    {
                        tmpPts.Remove(tmpPts[j]);
                    }
                }
            }
            return tmpPts;
        }








    }
}
