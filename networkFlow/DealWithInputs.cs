using DocumentFormat.OpenXml.Office2013.Word;
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
            return pts.Distinct(new Point3DCompare()).ToList();
        }

        public class Point3DCompare : IEqualityComparer<Point3d>
        {
            public bool Equals(Point3d x, Point3d y)
            {
                if (x == null || y == null)
                    return false;
                if (x == y)
                    return true;
                else
                    return false;
            }

            public int GetHashCode(Point3d obj)
            {
                if (obj == null)
                    return 0;
                else
                    return obj.GetHashCode();
            }
        }
    }
}
