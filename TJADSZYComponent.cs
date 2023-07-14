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
//using Rhino;


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
            pManager.AddBooleanParameter("run", "R", "是否运行", GH_ParamAccess.item, false);
            pManager.AddTextParameter("filePath", "file", "文件提前创建好写入GH电池里，如果不自己创建默认会在C盘根目录下创建名叫test的文件", GH_ParamAccess.item, "C:\\test.xlsx");
            pManager.AddTextParameter("startCellIndex", "cellInedex", "从哪个单元开始，默认为A1", GH_ParamAccess.item, "A1");
            pManager.AddTextParameter("titleNames", "titles", "大的分项名称", GH_ParamAccess.list);
            pManager.AddTextParameter("workSheetNames", "wsNames", "不同区域名称，也是分worksheet的名称，需跟工经确认好", GH_ParamAccess.list);
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            //pManager.AddTextParameter("test", "test", "test", GH_ParamAccess.item);
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
            #endregion

            #region make column widths and row heights
            List<int> columnWidth=new List<int>();

            columnWidth.Add(12);
            columnWidth.Add(20);
            columnWidth.Add(30);
            columnWidth.Add(40);
            columnWidth.Add(12);

            int titleRowHeight = 40;
            int otherRowHeight = 17;
            #endregion

            #region get all the matches from layer names
            string layerInfo = "";
            #endregion

            string tmpCellName;

            if (run)
            {
                #region get all layer informations
                foreach (Rhino.DocObjects.Layer layer in RhinoDoc.ActiveDoc.Layers)
                {
                    layerInfo += layer.FullPath;
                    layerInfo += "\r\n";
                }
                #endregion

                using (var workbook = new XLWorkbook())
                {
                    workbook.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    for (int i = 0; i < wsNames.Count; i++)
                    {
                        string pat = string.Format(@"000000算量::{0}::(.*)::(.*)::(.*)::(.*):(.*)", wsNames[i]);
                        workbook.Worksheets.Add(wsNames[i]);

                        #region fill in the titles
                        //make this title row bigger
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
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Value = "修改日期:" + "  " + DateTime.Now.ToString();
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).Style.Fill.SetBackgroundColor(XLColor.YellowMunsell);
                        workbook.Worksheet(wsNames[i]).Cell(tmpCellName).WorksheetColumn().Width = 60;
                        #endregion

                        int matchIndex = 1;
                        foreach (Match match in Regex.Matches(layerInfo, pat))
                        {
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

                            //make others row height
                            workbook.Worksheet(wsNames[i]).Cell(tmpCellName).WorksheetRow().Height = otherRowHeight;

                            matchIndex++;
                        }
                    }
                    workbook.SaveAs(filePath);
                }
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