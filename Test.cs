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
              "TJADSZY", "Test")
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
                double sumArea = 0;
                foreach (RhinoObject robj in tmpLayObjs)
                {
                    if (robj != null)
                    {
                        //if (robj.ObjectType == ObjectType.Brep || robj.ObjectType == ObjectType.Surface || robj.ObjectType == ObjectType.Extrusion)
                        //{
                        //    sumArea += AreaMassProperties.Compute(robj.Geometry as Brep).Area;
                        //}
                        if (robj.ObjectType == ObjectType.Brep)
                        {
                            sumArea += AreaMassProperties.Compute(robj.Geometry as Brep).Area;
                        }
                        else if (robj.ObjectType == ObjectType.Mesh)
                        {
                            sumArea += AreaMassProperties.Compute(robj.Geometry as Mesh).Area;
                        }
                        else
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                            //return;
                        }
                    }
                    else
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                        //return;
                    }
                }


                //layName = "curve";
                //tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(layName);
                //double sumLen = 0;
                //foreach (RhinoObject robj in tmpLayObjs)
                //{
                //    if (robj != null)
                //    {
                //        if (robj.ObjectType == ObjectType.Curve)
                //        {
                //            sumArea += LengthMassProperties.Compute(robj.Geometry as Curve).Length;
                //        }
                //        else
                //        {
                //            AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                //        }
                //    }
                //    else
                //    {
                //        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                //    }
                //}

                //layName = "others";
                //tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(layName);
                //double sumNum = 0;
                //foreach (RhinoObject robj in tmpLayObjs)
                //{
                //    if (robj != null)
                //    {
                //        sumNum++;
                //    }
                //    else
                //    {
                //        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                //    }
                //}

                DA.SetData("Brep", sumArea);
                //DA.SetData("Curve", sumLen);
                //DA.SetData("Num", sumNum);
            }


        }

        public string layErrorMassage(string layName)
        {
            //return string.Format("该图层有问题:{}", layName);
            return "该图层有问题";
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