using System;
using System.Collections.Generic;

using Rhino;
using Rhino.DocObjects;
using Grasshopper.Kernel;
using Rhino.Geometry;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TJADSZY
{
    public class Test2 : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the Test2 class.
        /// </summary>
        public Test2()
          : base("Test2", "Nickname",
              "Description",
              "TJADSZY", "TestForUnits")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("run", "R", "111", GH_ParamAccess.item, false);
            pManager.AddTextParameter("layerName", "layer", "test", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("units", "U", "test", GH_ParamAccess.item);
            pManager.AddGeometryParameter("layerObjs", "objs", "test", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool run = false;
            DA.GetData("run", ref run);
            string layerFullName = "";
            DA.GetData("layerName", ref layerFullName);
            List<GeometryBase> layObjs = new List<GeometryBase>();

            if (run)
            {
                RhinoObject[] tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(getLayerByFullName(layerFullName));
                foreach (RhinoObject roj in tmpLayObjs)
                {
                    layObjs.Add(roj.Geometry);
                }    
            }

            DA.SetDataList("layerObjs", layObjs);

        }

        public string layErrorMassage(string layName)
        {
            return string.Format("该图层有问题: {0}", layName);
        }

        public Layer getLayerByFullName(string layerFullName)
        {
            if (RhinoDoc.ActiveDoc.Layers.FindByFullPath(layerFullName, -1) != -1)
            {
                return RhinoDoc.ActiveDoc.Layers[RhinoDoc.ActiveDoc.Layers.FindByFullPath(layerFullName, -1)];
            }
            else
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layerFullName));
                return null;
            }
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
            get { return new Guid("B05F3F88-E44E-4B3B-8ED2-C6DB66F4134F"); }
        }
    }
}