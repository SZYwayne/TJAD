using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

using Grasshopper.Kernel;
using Rhino.Geometry;
using DocumentFormat.OpenXml.EMMA;
using System.Drawing;
using Aliyun.OSS;
using System.IO;

namespace TJADSZY.ai
{
    public class Configuration : GH_Component
    {


        /// <summary>
        /// Initializes a new instance of the Configuration class.
        /// </summary>
        public Configuration()
          : base("Configuration", "Config",
              "Description",
              "TJADSZY", "aiRender")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("input_active_or_not", "run", "active or not", GH_ParamAccess.item, false);
            pManager.AddTextParameter("input_base_model", "model", "choose your base model", GH_ParamAccess.item);
            pManager.AddTextParameter("input_prompt", "prompt", "enter your prompt", GH_ParamAccess.item);
            pManager.AddTextParameter("input_negative_prompt", "negPrompt", "enter your negative prompt", GH_ParamAccess.item);
            pManager.AddTextParameter("input_lora", "lora", "select lora models", GH_ParamAccess.list);
            pManager.AddIntegerParameter("input_image_number", "num", "how many images a set", GH_ParamAccess.item);
            pManager.AddIntegerParameter("input_image_width", "width", "image width", GH_ParamAccess.item);
            pManager.AddIntegerParameter("input_image_height", "height", "image height", GH_ParamAccess.item);
            pManager.AddTextParameter("input_nickname", "nickname", "enter your nickname to make sure your uploaded file doesnt overlap with others", GH_ParamAccess.item, "");
            pManager.AddBooleanParameter("input_controlnet_or_not", "controlnet", "use an image as input or not", GH_ParamAccess.item, false);
            pManager.AddTextParameter("input_namedView", "view", "view name", GH_ParamAccess.item, "");
            pManager.AddTextParameter("input_file_path", "file", "file path to hold your screenshot", GH_ParamAccess.item, "");
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("output_request_body", "body", "output the request body", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            
            bool active = false;
            if (!DA.GetData("input_active_or_not", ref active)) return;
            if (!active) { this.Message = "Innactive"; return; }

            this.Message = "Merging";

            string baseModel = "";
            if (!DA.GetData("input_base_model", ref baseModel)) return;
            if (baseModel == "") { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "baseModel has no input"); }

            string prompt = "";
            if (!DA.GetData("input_prompt", ref prompt)) return;
            prompt = prompt + "ARCHITECTURE,8k, realistic, 4k render, inspired by forster, inspired by SOM KPF,Best quality, masterpiece, MIR style";
            string negativePrompt = "";
            if (!DA.GetData("input_negative_prompt", ref negativePrompt)) return;
            negativePrompt = negativePrompt + "(worst quality, low quality:1.4), 2 faces, cropped image, out of frame, deformed hands, twisted fingers, double image, malformed hands, extra limb, poorly drawn hands, missing limb, disfigured, cut-off, grain, low-res, deformed, blurry, poorly drawn face, mutation, floating limbs, disconnected limbs, long body, disgusting, mutilated, extra fingers, duplicate artifacts, morbid, gross proportions, missing arms, mutated hands, malformed, ugly, tiling, poorly drawn hands, poorly drawn feet, poorly drawn face, out of frame, extra limbs, disfigured, deformed, body out of frame, bad anatomy, watermark, signature, cut off, low contrast, underexposed, overexposed, bad art, beginner, amateur, distorted face, blurry, draft, grainy, large scale, city scale, urban scale, people view, wierd picture, buildings in the middle";

            List<string> lora = new List<string>();
            if (!DA.GetDataList<string>("input_lora", lora)) return;
            for (int i=0;i<lora.Count; i++)
            {
                prompt += lora[i];
            }

            int batchSize = 1;
            if (!DA.GetData("input_image_number", ref batchSize)) return;
            if (batchSize <= 0) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "image number cant be less than 1"); return; }
            int width = 50;
            if (!DA.GetData("input_image_width", ref width)) return;
            if (width <= 100) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "image width cant be less than 100"); return; }
            int height = 50;
            if (!DA.GetData("input_image_height", ref height)) return;
            if (height <= 100) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "image height cant be less than 100"); }

            bool controlnet = false;
            if (!DA.GetData("input_controlnet_or_not", ref controlnet)) return;

            string nickName = "";
            if (!DA.GetData("input_nickname", ref nickName)) return;
            if (nickName == "" && controlnet) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "nickName has no input"); return; }
            nickName = "baseImg/" + nickName + ".png";

            string rawBody = @"{
                ""stable_diffusion_model"": """",
                ""denoising_strength"": 0,
                ""prompt"": """", 
                ""negative_prompt"":"""",
                ""seed"": -1,
                ""batch_size"": 5, 
                ""n_iter"": 1, 
                ""steps"": 50, 
                ""cfg_scale"": 7,
                ""width"": 700, 
                ""height"": 512,
                ""restore_faces"": false,
                ""tiling"": false,
                ""sampler_index"": ""Euler"",
                ""alwayson_scripts"": {
                ""controlnet"": {
                    ""args"": [
                            {
                                ""image"":""baseImg/test.png"", //支持传输base64和oss对应图片path(png/jpg/jpeg)
                                ""enabled"":true,
                                ""module"":""canny"",
                                ""model"":""control_v11p_sd15_canny_fp16"",
                                ""weight"":1,
                                ""resize_mode"":""Crop and Resize"",
                                ""low_vram"":false,
                                ""processor_res"":512,
                                ""threshold_a"":100,
                                ""threshold_b"":200,
                                ""guidance_start"":0,
                                ""guidance_end"":1,
                                ""pixel_perfect"":true,
                                ""control_mode"":0,
                                ""input_mode"":""simple""
                            }
                        ]
                    }
                }
            }";

            JObject tmpJsonObj = JObject.Parse(rawBody);

            tmpJsonObj["stable_diffusion_model"] = baseModel;
            tmpJsonObj["prompt"] = prompt;
            tmpJsonObj["negative_prompt"] = negativePrompt;
            tmpJsonObj["batch_size"] = batchSize;
            tmpJsonObj["width"] = width;
            tmpJsonObj["height"] = height;
            tmpJsonObj["alwayson_scripts"]["controlnet"]["args"][0]["image"] = nickName;
            tmpJsonObj["alwayson_scripts"]["controlnet"]["args"][0]["enabled"] = controlnet;

            if(controlnet)
            {
                this.Message = "Screenshotting";
                string namedView = "";
                if (!DA.GetData("input_namedView", ref namedView)) return;
                if (namedView == "" || namedView == null) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "namedView has no input"); return; }
                string filePath = "";
                if (!DA.GetData("input_file_path", ref filePath)) return;
                if (filePath == "" || filePath == null) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "filePath has no input"); return; }


                Rhino.RhinoDoc doc = Rhino.RhinoDoc.ActiveDoc;
                Rhino.Display.RhinoView view = doc.Views.Find(namedView, false);
                if (view == null) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "cant find that namedView, please check your view input"); return; }

                Bitmap image = view.CaptureToBitmap();
                if (image == null) { AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "cant take the screenshot, try in rhino see what's wrong"); return; }
                image.Save(filePath);
                image.Dispose();

                this.Message = "Uploading";
                string accessKeyId = "";
                string accessKeyS = "";
                string endpoint = "oss-cn-hangzhou.aliyuncs.com";
                string bucketName = "fc-sd-17ad442mg";
                string objectName = nickName;

                var ossClient = new OssClient(endpoint, accessKeyId, accessKeyS);
                var fileStream = new FileStream(filePath, FileMode.Open);
                PutObjectResult putObjectResult = ossClient.PutObject(bucketName, objectName, fileStream);
                fileStream.Close();
            }


            this.Message = "Done";
            string body = tmpJsonObj.ToString();

            DA.SetData("output_request_body", body);

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
            get { return new Guid("87D24459-1363-4052-8F41-5A48CBCF0045"); }
        }
    }
}