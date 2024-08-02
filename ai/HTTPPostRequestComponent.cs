using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Grasshopper.Kernel;
using Rhino;
using Rhino.Geometry;

namespace TJADSZY.ai
{
    public class HTTPPostRequestComponent : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the HTTPPostRequestComponent class.
        /// </summary>
        /// 


        private string _response = "";
        private bool _shouldExpire = false;

        public HTTPPostRequestComponent()
          : base("HTTPPostRequestComponent", "post",
              "Description",
              "TJADSZY", "aiRender")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Send", "run", "perform the request or not", GH_ParamAccess.item, false);
            pManager.AddTextParameter("URL", "url", "url for the request", GH_ParamAccess.item);
            pManager.AddTextParameter("BODY", "body", "request body", GH_ParamAccess.item);
            pManager.AddIntegerParameter("Timeout", "maxTime", "time for the request in ms. if longer thanthis, it will break and fail.(default if 10000)", GH_ParamAccess.item, 10000);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("outputResponse", "response", "request response", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            if (_shouldExpire)
            {
                DA.SetData("outputResponse", _response);
                _shouldExpire = false;
                return;
            }


            bool active = false;
            DA.GetData("Send", ref active);
            if (!active) { DA.SetData("Response", ""); return; }

            string url = "";
            if (!DA.GetData("URL", ref url)) return;
            string body = "";
            if (!DA.GetData("BODY", ref body)) return;
            int timeout = 0;
            if (!DA.GetData("Timeout", ref timeout)) return;

            if (url==null || url.Length==0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "empty url");
                return;
            }

            AsyncPOST(url, body, timeout);
        }

        protected override void ExpireDownStreamObjects()
        {
            if(_shouldExpire)
            {
                base.ExpireDownStreamObjects();
            }
        }

        private void AsyncPOST(
            string url,
            string body,
            int timeout)
        {

            Task.Run(() =>
            {
                try
                {
                    //deal with the request
                    byte[] data = Encoding.ASCII.GetBytes(body);

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    request.Method = "POST";
                    request.ContentType = "application/json";
                    request.ContentLength = data.Length;
                    request.Timeout = timeout;
                    request.Credentials = CredentialCache.DefaultCredentials;

                    using (var stream = request.GetRequestStream())
                    {
                        stream.Write(data, 0, data.Length);
                    }
                    var res = request.GetResponse();
                    _response = new StreamReader(res.GetResponseStream()).ReadToEnd();

                    _shouldExpire = true;
                    RhinoApp.InvokeOnUiThread((Action)delegate { ExpireSolution(true); });
                }
                catch (Exception ex)
                {
                    _response = ex.Message;
                    _shouldExpire = true;
                    RhinoApp.InvokeOnUiThread((Action)delegate { ExpireSolution(true); });
                    return;
                }
            });

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
            get { return new Guid("79F23F19-07C6-451F-8AD4-8307183C2988"); }
        }
    }
}