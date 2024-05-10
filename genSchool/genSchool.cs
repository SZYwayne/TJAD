using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace TJADSZY.genSchool
{
    public class genSchool : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the genSchool class.
        /// </summary>
        public genSchool()
          : base("genSchool", "GS",
              "Description",
              "SJTUSZY", "genSchool")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddNumberParameter("plotRatio", "PR", "容积率(默认为1.5)", GH_ParamAccess.item, 1.5);
            pManager.AddPointParameter("boundaryPoints", "nakedPts", "矩形的边界点", GH_ParamAccess.list);
            pManager.AddCurveParameter("playgroundS", "PGs", "最小操场", GH_ParamAccess.list);
            pManager.AddCurveParameter("playgroundM", "PGm", "中间操场", GH_ParamAccess.list);
            pManager.AddCurveParameter("playgroundL", "PGl", "最大操场", GH_ParamAccess.list);
            pManager.AddNumberParameter("dayFactor", "dayFactor", "日照系数", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddCurveParameter("outputPG", "PG", "playground", GH_ParamAccess.list);
            pManager.AddBoxParameter("outputTeaching", "teaching", "teachingBuildings", GH_ParamAccess.list);
            pManager.AddBoxParameter("outputOther", "other", "otherBuildings", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get inputs
            List<Curve> PGs = new List<Curve>();
            if (!DA.GetDataList("playgroundS", PGs)) return;
            List<Curve> PGm = new List<Curve>();
            if (!DA.GetDataList("playgroundM", PGm)) return;
            List<Curve> PGl = new List<Curve>();
            if (!DA.GetDataList("playgroundL", PGl)) return;
            List<Point3d> nakedPts = new List<Point3d>();
            if (!DA.GetDataList("boundaryPoints", nakedPts)) return;
            double dayFactor = 0;
            if (!DA.GetData("dayFactor", ref dayFactor)) return;
            double PR = 0;
            if (!DA.GetData("plotRatio", ref PR)) return;
            #endregion


            #region outputs

            DA.SetDataList("outputPG", (List<Curve>)utilities.allBuildings(PGs, PGm, PGl, nakedPts, PR, dayFactor)[0]);
            DA.SetDataList("outputTeaching", (List<Box>)utilities.allBuildings(PGs, PGm, PGl, nakedPts, PR, dayFactor)[1]);
            DA.SetDataList("outputOther", (List<Box>)utilities.allBuildings(PGs, PGm, PGl, nakedPts, PR, dayFactor)[2]);

            #endregion
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
            get { return new Guid("7540E496-43E0-4132-B136-93812B358097"); }
        }
    }
}