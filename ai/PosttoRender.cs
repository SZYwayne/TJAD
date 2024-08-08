using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Grasshopper.Kernel;
using Rhino;
using System.Text.RegularExpressions;

namespace TJADSZY.ai
{
    public class PosttoRender : GH_Component
    {
        private string _response = "";
        private bool _shouldExpire = false;
        private RequestState _currentState = RequestState.Off;

        /// <summary>
        /// Initializes a new instance of the PosttoRender class.
        /// </summary>
        public PosttoRender()
          : base("PosttoRender", "render",
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
            pManager.AddTextParameter("BODY", "body", "request body", GH_ParamAccess.item);
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
                switch (_currentState)
                {
                    case RequestState.Off:
                        this.Message = "Inactive";
                        DA.SetData("outputResponse", "");
                        _currentState = RequestState.Idle;
                        break;

                    case RequestState.Error:
                        this.Message = "ERROR";
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, _response);
                        _currentState = RequestState.Idle;
                        break;

                    case RequestState.Done:
                        this.Message = "Complete!!!";
                        DA.SetData("outputResponse", _response);
                        _currentState = RequestState.Idle;
                        break;
                }

                string pattern = @"\\u0026";
                string resultResponse = Regex.Replace(_response, pattern, "&");
                DA.SetData("outputResponse", resultResponse);
                _shouldExpire = false;
                return;
            }


            bool active = false;
            DA.GetData("Send", ref active);
            if (!active)
            {
                _currentState = RequestState.Off;
                _shouldExpire = true;
                ExpireSolution(true);
                return;
            }

            string body = "";
            if (!DA.GetData("BODY", ref body)) return;

            string url = "http://sd-692728--proxy.fcv3.1585809013229189.cn-hangzhou.fc.devsapp.net/txt2img";
            int timeout = 1200000;

            _currentState = RequestState.Requesting;
            this.Message = "Requesting...";

            AsyncPOST(url, body, timeout);
        }

        protected override void ExpireDownStreamObjects()
        {
            if (_shouldExpire)
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

                    _currentState = RequestState.Done;

                    _shouldExpire = true;
                    RhinoApp.InvokeOnUiThread((Action)delegate { ExpireSolution(true); });
                }
                catch (Exception ex)
                {
                    _response = ex.Message;
                    _shouldExpire = true;

                    _currentState = RequestState.Error;

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
            get { return new Guid("7D328C92-692E-4092-9688-8B5EF174D82F"); }
        }
    }
}