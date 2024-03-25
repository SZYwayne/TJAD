using System;
using System.Collections.Generic;
using System.Linq;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Rhino.NodeInCode;

namespace TJADSZY.package.Handles
{
    public class HandRail : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public HandRail()
          : base("HandRail", "HR",
              "添加细部构件",
              "TJADSZY", "Package")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddCurveParameter("inputRail", "rail", "需要做构件的轨道线", GH_ParamAccess.list);
            pManager.AddIntegerParameter("inputCount", "count", "分多少个点", GH_ParamAccess.item, 3);
            pManager.AddNumberParameter("inputHeight", "height", "构件高度默认为1", GH_ParamAccess.item, 1);
            pManager.AddCurveParameter("inputSection", "section", "构件截面,必须是封闭曲线", GH_ParamAccess.item);
            pManager.AddPointParameter("inputLocation", "location", "截面相对位置", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddBrepParameter("outputBrep", "details", "生成构件们", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get inputs
            List<Curve> rails = new List<Curve>();
            if (!DA.GetDataList("inputRail", rails)) return;
            int count = 0;
            if (!DA.GetData("inputCount", ref count)) return;
            double height = 0;
            if (!DA.GetData("inputHeight", ref height)) return;
            Curve section = null;
            if (!DA.GetData("inputSection", ref section)) section = new Rectangle3d(Plane.WorldXY, 1, 1).ToNurbsCurve();
            Point3d location = new Point3d();
            if (!DA.GetData("inputLocation", ref location)) return;
            #endregion

            List<Brep> outputs = new List<Brep>();
            Plane origin = new Plane(location, Vector3d.XAxis, Vector3d.YAxis);

            if (rails.Count == 1)
            {
                List<Curve> sections = getAllSectionsInPosition(rails[0], section, count, origin);
                var funExtrude = Components.FindComponent("Extrude").Delegate as dynamic;
                foreach(Curve c in sections)
                {
                    Brep tmp = funExtrude(c, Vector3d.ZAxis * height)[0];
                    outputs.Add(tmp);
                }
            }
            else if (rails.Count == 2)
            {
                List<Curve> sections0 = getAllSectionsInPosition(rails[0], section, count, origin);
                List<Curve> sections1 = getAllSectionsInPosition(rails[1], section, count, origin);
                var funLoft = Components.FindComponent("Loft").Delegate as dynamic;
                for (int i = 0; i < sections0.Count; i++)
                {
                    List<Curve> tmp = new List<Curve>();
                    tmp.Add(sections0[i]);
                    tmp.Add(sections1[i]);
                    outputs.Add(Brep.CreateFromLoft(tmp, Point3d.Unset, Point3d.Unset, LoftType.Normal, false)[0]);
                }
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "输入的线数量不对");
            }
            DA.SetDataList("outputBrep", outputs);
        }

        private List<Curve> getAllSectionsInPosition(Curve rail, Curve section, int Num, Plane origin)
        {
            var funDivideCrv = Components.FindComponent("DivideCurve").Delegate as dynamic;
            var funOrient = Components.FindComponent("Orient").Delegate as dynamic;

            var tmpDividedPts = funDivideCrv(rail, Num, false)[0] as IList<object>;
            var tmpTangents = funDivideCrv(rail, Num, false)[1] as IList<object>;


            List<Point3d> dividedPts = tmpDividedPts.OfType<Point3d>().ToList();
            List<Vector3d> tangents = tmpTangents.OfType<Vector3d>().ToList();
            List<Plane> targets = new List<Plane>();
            List<Curve> sections = new List<Curve>();

            for (int i = 0; i < dividedPts.Count; i++)
            {
                Plane tmp = new Plane(dividedPts[i], Vector3d.CrossProduct(tangents[i], Vector3d.ZAxis), tangents[i]);
                targets.Add(tmp);
            }

            var tmpSections = funOrient(section, origin, targets)[0] as IList<object>;
            sections = tmpSections.OfType<Curve>().ToList();
            return sections;
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
            get { return new Guid("EF319AB9-2670-40F7-841E-D6E368017068"); }
        }
    }
}