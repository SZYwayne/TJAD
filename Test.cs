using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using Rhino.DocObjects;
using Rhino;
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TJADSZY
{
    public class Test : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the Test class.
        /// </summary>
        public Test()
          : base("Test", "Nickname",
              "Description",
              "TJADSZY", "TestForlayerObjects")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("run", "R", "111", GH_ParamAccess.item, false);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddNumberParameter("Brep", "AreaBrep", "test", GH_ParamAccess.item);
            pManager.AddNumberParameter("Curve", "LenCurve", "test", GH_ParamAccess.item);
            pManager.AddNumberParameter("Num", "OtherNum", "test", GH_ParamAccess.item);
            pManager.AddGenericParameter("test", "test", "test", GH_ParamAccess.item) ;
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool run = false;
            DA.GetData("run", ref run);

            if (run)
            {
                string layName = "brep";
                RhinoObject[] tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(layName);
                double tmpSumArea = sumArea(layName, tmpLayObjs);

                layName = "curve";
                tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(layName);
                double tmpSumLen = sumLen(layName, tmpLayObjs);

                layName = "others";
                tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(layName);
                int tmpSumNum = sumNum(layName, tmpLayObjs);

                DA.SetData("Brep", tmpSumArea);
                DA.SetData("Curve", tmpSumLen);
                DA.SetData("Num", tmpSumNum);
            }
        }

        public string layErrorMassage(string layName)
        {
            return string.Format("该图层有问题: {0}", layName);
        }

        public double sumArea(string layName, RhinoObject[] tmpLayObjs)
        {
            double sum = 0;
            foreach (RhinoObject robj in tmpLayObjs)
            {
                if (robj != null)
                {
                    if (robj.ObjectType == ObjectType.Brep)
                    {
                        sum += AreaMassProperties.Compute(robj.Geometry as Brep).Area;
                    }
                    else if (robj.ObjectType == ObjectType.Extrusion)
                    {
                        sum += AreaMassProperties.Compute(robj.Geometry as Extrusion).Area;
                    }
                    else if (robj.ObjectType == ObjectType.Surface)
                    {
                        sum += AreaMassProperties.Compute(robj.Geometry as Surface).Area;
                    }
                    else if (robj.ObjectType == ObjectType.Mesh)
                    {
                        sum += AreaMassProperties.Compute(robj.Geometry as Mesh).Area;
                    }
                    else
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                    }
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                }
            }
            return sum;
        }

        public double sumLen(string layName, RhinoObject[] tmpLayObjs)
        {
            double sum = 0;
            foreach (RhinoObject robj in tmpLayObjs)
            {
                if (robj != null)
                {
                    if (robj.ObjectType == ObjectType.Curve)
                    {
                        sum += LengthMassProperties.Compute(robj.Geometry as Curve).Length;
                    }
                    else
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                        //DA.SetData("test", robj.ObjectType);
                    }
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                }
            }
            return sum;
        }

        public int sumNum(string layName, RhinoObject[] tmpLayObjs)
        {
            int sum = 0;
            foreach (RhinoObject robj in tmpLayObjs)
            {
                if (robj != null)
                {
                    sum++;
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                }
            }
            return sum;
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
            get { return new Guid("EE5CFAAB-1DB0-4E7D-A405-6F9715C09C97"); }
        }
    }
}