using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Rhino.NodeInCode;
using Rhino.Display;
using Rhino;
using DocumentFormat.OpenXml.Spreadsheet;

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
            pManager.AddNumberParameter("inputCapacity", "capacities", "test", GH_ParamAccess.list);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddNumberParameter("outputMaxFlow", "maxFlow", "test", GH_ParamAccess.item);
            pManager.AddPlaneParameter("outputPlane", "Plane", "test", GH_ParamAccess.list);
            pManager.AddTextParameter("outputEdgeInfo", "edgeInfo", "test", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get variables from outside
            List<Curve> lines = new List<Curve>();
            if (!DA.GetDataList("inputEdges", lines)) return;
            List<double> capacities = new List<double>();
            if (!DA.GetDataList("inputCapacity", capacities)) return;

            string inputWarning = "输入的edge和capacity长度不一样";
            if (lines.Count !=  capacities.Count)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, inputWarning);
            }
            #endregion

            #region get all the nodes from the edges
            List<Point3d> allPts = new List<Point3d>();
            foreach(Curve crv in lines)
            {
                allPts.Add(crv.PointAtStart);
                allPts.Add(crv.PointAtEnd);
            }
            List<Point3d> nodes = DealWithInputs.noRepeatPts(allPts);
            #endregion

            #region begin & end point must be 0 and n-1
            int n = nodes.Count;
            int s = 0;
            int t = n - 1;
            NetworkFlowSolverBase solver = new FordFulkersonDfsSolver(n, s, t);
            for(int i = 0; i < lines.Count; i++)
            {
                solver.addEdge(nodes.FindIndex(a => a == lines[i].PointAtStart), nodes.FindIndex(a => a == lines[i].PointAtEnd), capacities[i]);
            }
            #endregion

            List<Edge>[] resultGraph = solver.getGraph();
            List<Plane> location = new List<Plane>();
            List<string> edgeInfos = new List<string>();
            Point3d tmpCen;
            Vector3d vx;
            Vector3d vy;
            double splitText = 5;
            foreach(List<Edge> edges in resultGraph)
            {
                foreach(Edge e in edges)
                {
                    tmpCen = (nodes[e.from] + nodes[e.to]) / 2;
                    vx = new Vector3d(nodes[e.to] - nodes[e.from]);
                    vy = Vector3d.CrossProduct(vx, Vector3d.ZAxis);
                    vy = vy / vy.Length;
                    if (e.isResidual())
                    {
                        continue;
                    }
                    else
                    {
                        location.Add(new Plane((tmpCen + vy * splitText), -1 * vx, vy));
                        edgeInfos.Add(e.edgeInfo());
                    }
                }
            }
            DA.SetData("outputMaxFlow", solver.getMaxFlow());
            DA.SetDataList("outputEdgeInfo", edgeInfos);
            DA.SetDataList("outputPlane", location);
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