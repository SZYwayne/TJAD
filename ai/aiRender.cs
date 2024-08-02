using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

using Grasshopper.Kernel;
using System.Drawing;
namespace TJADSZY.ai
{
    

    public class aiRender : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the aiRender class.
        /// </summary>
        /// 


        public aiRender()
          : base("aiRender", "AR",
              "Description",
              "TJADSZY", "aiRender")
        {
        }


        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        /// 





        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("prompt", "prompt", "文字描述（默认为realistic, 8k, hd, architectural rendering）", GH_ParamAccess.item, "realistic, 8k, hd, architectural rendering");
            pManager.AddTextParameter("file_path_to_hold_the_image", "filePath", "图片地址", GH_ParamAccess.item,"./11.png");
            pManager.AddTextParameter("named_view_to_render", "namedView", "需要出图的视角", GH_ParamAccess.item, "Perspective");
            pManager.AddBooleanParameter("toggle_to_run", "run", "点击运行", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("outputImage", "imageURL", "imageURL", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        /// 


        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get inputs
            string promot = "";
            if (!DA.GetData("prompt", ref promot)) return;
            string filePath="";
            if (!DA.GetData("file_path_to_hold_the_image", ref filePath)) return;
            string namedView = "";
            if (!DA.GetData("named_view_to_render", ref namedView)) return;
            bool run =false;
            if (!DA.GetData("toggle_to_run", ref run)) return;
            #endregion


            //Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
            //Rhino.Display.RhinoView view = doc.Views.Find(namedView, false);
            //if (view == null)
            //    return;
            //Bitmap image = view.CaptureToBitmap();
            //if (image == null)
            //    return;
            //image.Save(filePath);
            //image.Dispose();


            



            if (run)
            {

            }
            #region outputs
            DA.SetData("outputImage", filePath);
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
            get { return new Guid("5515ADF0-E7AD-4E28-8D56-CFB63473C6D5"); }
        }
    }
}