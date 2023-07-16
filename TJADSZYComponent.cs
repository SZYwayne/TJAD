using ClosedXML.Excel;
using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;
using System.Text.RegularExpressions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using Rhino;
using Rhino.DocObjects;
using DocumentFormat.OpenXml.Wordprocessing;


namespace TJADSZY
{
    public class TJADSZYComponent : GH_Component
    {
        /// <summary>
        /// Each implementation of GH_Component must provide a public 
        /// constructor without any arguments.
        /// Category represents the Tab in which the component will appear, 
        /// Subcategory the panel. If you use non-existing tab or panel names, 
        /// new tabs/panels will automatically be created.
        /// </summary>
        public TJADSZYComponent()
          : base("TJADSZY", "Nickname",
            "Description",
            "TJADSZY", "Base")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("run", "R", "�Ƿ�����", GH_ParamAccess.item, false);
            pManager.AddTextParameter("filePath", "file", "�ļ���ǰ������д��GH����������Լ�����Ĭ�ϻ���C�̸�Ŀ¼�´�������test���ļ�", GH_ParamAccess.item, "C:\\test.xlsx");
            pManager.AddTextParameter("startCellIndex", "cellInedex", "���ĸ���Ԫ��ʼ��Ĭ��ΪA1", GH_ParamAccess.item, "A1");
            pManager.AddTextParameter("titleNames", "titles", "��ķ�������", GH_ParamAccess.list);
            pManager.AddTextParameter("workSheetNames", "wsNames", "��ͬ�������ƣ�Ҳ�Ƿ�worksheet�����ƣ���ͼ��ҲϢϢ��أ����ﲻд����Ƕ�ȡ������Ӧͼ���", GH_ParamAccess.list);

            pManager.AddNumberParameter("modForBLQM", "modGlassWall", "������ǽ��ķ���", GH_ParamAccess.item, 1);
            pManager.AddNumberParameter("modForQTQM", "modOtherWall", "������ǽ��ķ���", GH_ParamAccess.item, 1);
            pManager.AddNumberParameter("modForDD", "modCeiling", "�������ķ���", GH_ParamAccess.item, 1);
            pManager.AddNumberParameter("modForDDFB", "modCeilingSide", "���������ߵķ���", GH_ParamAccess.item, 1);
            pManager.AddNumberParameter("modForDM", "modFloor", "������ķ���", GH_ParamAccess.item, 1);
            pManager.AddNumberParameter("modForLG", "modHandles", "�����˵ķ���", GH_ParamAccess.item, 1);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("test", "test", "test", GH_ParamAccess.item);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object can be used to retrieve data from input parameters and 
        /// to store data in output parameters.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            #region get variables from outside
            bool run = false;
            DA.GetData(0, ref run);
            string filePath = "C:\\test.xlsx";
            DA.GetData(1, ref filePath);
            string cellName = "A1";
            DA.GetData(2, ref cellName);
            List<string> titleNames = new List<string>();
            DA.GetDataList(3, titleNames);
            List<string> wsNames = new List<string>();
            DA.GetDataList(4, wsNames);
            double modBLQM = 1;
            DA.GetData("modForBLQM", ref modBLQM);
            double modQTQM = 1;
            DA.GetData("modForQTQM", ref modQTQM);
            double modDD = 1;
            DA.GetData("modForDD", ref modDD);
            double modDDFB = 1;
            DA.GetData("modForDDFB", ref modDDFB);
            double modDM = 1;
            DA.GetData("modForDM", ref modDM);
            double modLG = 1;
            DA.GetData("modForLG", ref modLG);
            #endregion

            #region make column widths and row heights
            List<int> columnWidth=new List<int>();

            columnWidth.Add(12);
            columnWidth.Add(20);
            columnWidth.Add(30);
            columnWidth.Add(40);
            columnWidth.Add(12);
            columnWidth.Add(20);
            columnWidth.Add(20);

            int titleRowHeight = 40;
            int otherRowHeight = 17;
            #endregion

            string layerInfo = "";
            string tmpCellName;
            string wainingUnits = "��ģ�͵�λ�Ǻ��׻���";

            if (run)
            {
                #region make sure the unit is right
                double modUnits = 1;
                if (RhinoDoc.ActiveDoc.GetUnitSystemName(true, false, false, false) == "����" || RhinoDoc.ActiveDoc.GetUnitSystemName(true, false, false, false) == "millimeters")
                {
                    modUnits = 0.001;
                }
                else if (RhinoDoc.ActiveDoc.GetUnitSystemName(true, false, false, false) == "��" || RhinoDoc.ActiveDoc.GetUnitSystemName(true, false, false, false) == "meters")
                {
                    modUnits = 1;
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, wainingUnits);
                }
                #endregion

                #region get all layer informations
                foreach (Rhino.DocObjects.Layer layer in RhinoDoc.ActiveDoc.Layers)
                {
                    layerInfo += layer.FullPath;
                    layerInfo += "\n";
                }
                #endregion

                using (var workbook = new XLWorkbook())
                {
                    workbook.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    for (int i = 0; i < wsNames.Count; i++)
                    {
                        string pat = string.Format(@"000000����::{0}::(.*)::(.*)::(.*)::(.*):(.*)", wsNames[i]);
                        workbook.Worksheets.Add(wsNames[i]);

                        #region fill in the titles
                        tmpCellName = Convert.ToChar(cellName[0]).ToString() + (Convert.ToInt32(cellName[1].ToString())).ToString();
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).WorksheetRow().Height = titleRowHeight;
                        for (int j = 0; j < titleNames.Count; j++)
                        {
                            tmpCellName = Convert.ToChar(cellName[0] + j).ToString() + (Convert.ToInt32(cellName[1].ToString())).ToString();
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = titleNames[j];
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).WorksheetColumn().Width = columnWidth[j];
                        }
                        #endregion

                        #region current time sample
                        tmpCellName = Convert.ToChar(cellName[0] + titleNames.Count).ToString() + (Convert.ToInt32(cellName[1].ToString())).ToString();
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = "�޸�����:" + "  " + DateTime.Now.ToString();
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Style.Fill.SetBackgroundColor(XLColor.YellowMunsell);
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).WorksheetColumn().Width = 60;
                        #endregion

                        int matchIndex = 1;
                        foreach (Match match in Regex.Matches(layerInfo, pat))
                        {
                            string layName = match.Groups[0].Value;

                            tmpCellName = Convert.ToChar(cellName[0] + 0).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = match.Groups[1].Value;

                            tmpCellName = Convert.ToChar(cellName[0] + 1).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = match.Groups[2].Value;

                            tmpCellName = Convert.ToChar(cellName[0] + 2).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = match.Groups[3].Value;

                            tmpCellName = Convert.ToChar(cellName[0] + 3).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = match.Groups[4].Value;

                            tmpCellName = Convert.ToChar(cellName[0] + 4).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = match.Groups[5].Value;

                            RhinoObject[] tmpLayObjs = RhinoDoc.ActiveDoc.Objects.FindByLayer(getLayerByFullName(layName));
                            //if (tmpLayObjs == null)
                            //{
                            //    //AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "������");
                            //    DA.SetData("test", layName);
                            //}

                            if (match.Groups[5].Value == "ƽ����")
                            {
                                double sumA = 0;
                                double modSumA = 0;
                                sumA = sumArea(layName, tmpLayObjs) * modUnits * modUnits;
                                modSumA = sumA;
                                if (match.Groups[3].Value == "����ǽ��")
                                {
                                    modSumA *= modBLQM;
                                }
                                else if (match.Groups[3].Value == "����ǽ��")
                                {
                                    modSumA *= modQTQM;
                                }
                                else if (match.Groups[3].Value == "����")
                                {
                                    modSumA *= modDD;
                                }
                                else if (match.Groups[3].Value == "��������")
                                {
                                    modSumA *= modDDFB;
                                }
                                else if (match.Groups[3].Value == "����")
                                {
                                    modSumA *= modDM;
                                }
                                else
                                {
                                    modSumA = sumA;
                                }
                                tmpCellName = Convert.ToChar(cellName[0] + 5).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                                workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = modSumA;
                                tmpCellName = Convert.ToChar(cellName[0] + 6).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                                workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = sumA;
                            }
                            else if (match.Groups[5].Value == "��")
                            {
                                double sumL = 0;
                                double modSumL = 0;
                                sumL = sumLen(layName, tmpLayObjs) * modUnits;
                                modSumL = sumL;
                                if (match.Groups[3].Value == "����")
                                {
                                    modSumL *= modLG;
                                }
                                else
                                {
                                    modSumL = sumL;
                                }
                                tmpCellName = Convert.ToChar(cellName[0] + 5).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                                workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = modSumL;
                                tmpCellName = Convert.ToChar(cellName[0] + 6).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                                workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = sumL;
                            }
                            else if (match.Groups[5].Value == "��")
                            {
                                int sumN = 0;
                                sumN = sumNum(layName, tmpLayObjs);
                                
                                tmpCellName = Convert.ToChar(cellName[0] + 5).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                                workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = sumN;
                                tmpCellName = Convert.ToChar(cellName[0] + 6).ToString() + (Convert.ToInt32(cellName[1].ToString()) + matchIndex).ToString();
                                workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = sumN;
                            }
                            else
                            {
                                //AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "������");
                                DA.SetData("test", match.Groups[5].Value);
                            }

                            //make others row height
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).WorksheetRow().Height = otherRowHeight;

                            matchIndex++;
                        }
                    }
                    workbook.SaveAs(filePath);
                }
            }
        }

        public string layErrorMassage(string layName)
        {
            return string.Format("��ͼ��������: {0}", layName);
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
                        if (AreaMassProperties.Compute(robj.Geometry as Extrusion) != null)
                        {
                            sum += AreaMassProperties.Compute(robj.Geometry as Extrusion).Area;
                        }
                        else
                        {
                            AddRuntimeMessage(GH_RuntimeMessageLevel.Error, layErrorMassage(layName));
                        }
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
        /// Provides an Icon for every component that will be visible in the User Interface.
        /// Icons need to be 24x24 pixels.
        /// You can add image files to your project resources and access them like this:
        /// return Resources.IconForThisComponent;
        /// </summary>
        protected override System.Drawing.Bitmap Icon => null;

        /// <summary>
        /// Each component must have a unique Guid to identify it. 
        /// It is vital this Guid doesn't change otherwise old ghx files 
        /// that use the old ID will partially fail during loading.
        /// </summary>
        public override Guid ComponentGuid => new Guid("37288779-ca15-43d6-a888-ad730739e58f");
    }
}