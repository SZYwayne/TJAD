using System;
using System.Collections.Generic;
using System.Linq;
using Grasshopper.Kernel;
using Rhino.Geometry;

namespace TJADSZY.package.NewGraph
{
    public class GraphAsWish : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GraphAsWish()
          : base("GraphAsWish", "GAW",
              "test",
              "TJADSZY", "Package")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddCurveParameter("inputCrv", "Crv", "用来操作的单根曲线，必须是原始XYZ下画出的曲线", GH_ParamAccess.item);
            pManager.AddIntegerParameter("inputNumber", "Count", "需要多少个数据", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddNumberParameter("outputNumbers", "data", "契合输入曲线的0到1之间的一组数据", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get all the inputs
            int Num = 0; 
            if (!DA.GetData("inputNumber", ref Num)) return;
            Curve crv = null;
            if (!DA.GetData("inputCrv", ref crv)) return;
            #endregion

            List<double> outputValues = new List<double>();
            List<double> yValues = new List<double>();

            Point3d[] dividedPts;
            crv.DivideByCount(Num, true, out dividedPts);

            foreach (Point3d point in dividedPts)
            {
                yValues.Add(point.Y);
            }

            double range = yValues.Max() - yValues.Min();

            foreach (Point3d point in dividedPts)
            {
                outputValues.Add((point.Y - yValues.Min()) / range);
            }
            DA.SetDataList("outputNumbers", outputValues);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("002E6409-5AED-4166-8E83-15D04DBF4540"); }
        }
    }
}