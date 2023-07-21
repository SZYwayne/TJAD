using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace TJADSZY.networkFlow
{
    public class networkFlow : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public networkFlow()
          : base("networkFLow", "NF",
              "Description",
              "TJADSZY", "networkFlow")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddCurveParameter("inputEdges", "edges", "test", GH_ParamAccess.list);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddPointParameter("outputPts", "nodes", "test", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get variables from outside
            List<Curve> edges = new List<Curve>();
            if (!DA.GetData("inputEdges", ref edges)) return;
            #endregion

            List<Point3d> allPts = new List<Point3d>();
            foreach(Curve crv in edges)
            {
                allPts.Add(crv.PointAtStart);
                allPts.Add(crv.PointAtEnd);
            }

            List<Point3d> nodes = DealWithInputs.noRepeatPts(allPts);


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
            get { return new Guid("2EF9F2ED-A51B-41D5-A5D1-0CA6F5A5755B"); }
        }
    }
}