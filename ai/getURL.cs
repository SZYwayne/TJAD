using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;

namespace TJADSZY.ai
{
    public class getURL : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the getURL class.
        /// </summary>
        public getURL()
          : base("getURL", "getUrl",
              "Description",
              "TJADSZY", "aiRender")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("FulUrl", "response", "full response input", GH_ParamAccess.item);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("outputUrls", "imageUrl", "all url for image sampler", GH_ParamAccess.list);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string fullUrl = "";
            if (!DA.GetData("FulUrl", ref fullUrl)) return;

            List<string> urls = utilities_ai.ExtractUrls(fullUrl);
            DA.SetDataList("outputUrls", urls);
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
            get { return new Guid("B4BFD2C3-08A2-4BCB-A53D-80D503ED17CC"); }
        }
    }
}