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
            //string layerInfo = "LINE\r\nSZY\r\n000000算量\r\n000000算量::定位点以防万一\r\n000000算量::站房部分\r\n000000算量::站房部分::-9.0\r\n000000算量::站房部分::-9.0::出站厅\r\n000000算量::站房部分::-9.0::出站厅::其他墙面\r\n000000算量::站房部分::-9.0::出站厅::其他墙面::<N1>干挂花岗岩内墙面:平方米\r\n000000算量::站房部分::-9.0::出站厅::吊顶\r\n000000算量::站房部分::-9.0::出站厅::吊顶::<C2>铝单板吊顶:平方米\r\n000000算量::站房部分::-9.0::出站厅::吊顶翻边\r\n000000算量::站房部分::-9.0::出站厅::吊顶翻边::<C2>铝单板吊顶:平方米\r\n000000算量::站房部分::-9.0::出站厅::地面\r\n000000算量::站房部分::-9.0::出站厅::地面::<F1C>花岗岩楼面三:平方米\r\n000000算量::站房部分::-9.0::出站厅::其他工程\r\n000000算量::站房部分::-9.0::出站厅::其他工程::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::站房部分::0.0\r\n000000算量::站房部分::0.0::旅服中庭\r\n000000算量::站房部分::0.0::旅服中庭::其他墙面\r\n000000算量::站房部分::0.0::旅服中庭::其他墙面::<W1B>干挂铝板柱面:平方米\r\n000000算量::站房部分::0.0::旅服中庭::吊顶\r\n000000算量::站房部分::0.0::旅服中庭::吊顶::<C1>穿孔铝板 造型铝板吊顶:平方米\r\n000000算量::站房部分::0.0::旅服中庭::吊顶翻边\r\n000000算量::站房部分::0.0::旅服中庭::吊顶翻边::<C1>穿孔铝板 造型铝板吊顶:平方米\r\n000000算量::站房部分::0.0::旅服中庭::地面\r\n000000算量::站房部分::0.0::旅服中庭::地面::<F1C>花岗岩楼面三:平方米\r\n000000算量::站房部分::0.0::旅服中庭::栏杆\r\n000000算量::站房部分::0.0::旅服中庭::栏杆::<Q8> 1.3m 8+1.52PVB+8:米\r\n000000算量::站房部分::0.0::旅服中庭::其他工程\r\n000000算量::站房部分::0.0::旅服中庭::其他工程::1.5m挡烟垂壁:米\r\n000000算量::站房部分::0.0::行政贵宾\r\n000000算量::站房部分::0.0::行政贵宾::其他墙面\r\n000000算量::站房部分::0.0::行政贵宾::其他墙面::<N4>干挂大理石墙面:平方米\r\n000000算量::站房部分::0.0::行政贵宾::吊顶\r\n000000算量::站房部分::0.0::行政贵宾::吊顶::<C9>造型石膏板吊顶(有改动):平方米\r\n000000算量::站房部分::0.0::行政贵宾::地面\r\n000000算量::站房部分::0.0::行政贵宾::地面::<F7A>大理石楼面:平方米\r\n000000算量::站房部分::0.0::行政贵宾::其他工程\r\n000000算量::站房部分::0.0::行政贵宾::其他工程::150mm高踢脚:米\r\n000000算量::站房部分::0.0::行政贵宾::其他工程::沙发(有改动):个\r\n000000算量::站房部分::0.0::行政贵宾::其他工程::贵宾软装(有改动):平方米\r\n000000算量::站房部分::0.0::行政贵宾::其他工程::地毯(有改动):平方米\r\n000000算量::站房部分::9.9\r\n000000算量::站房部分::9.9::候车大厅\r\n000000算量::站房部分::9.9::候车大厅::其他墙面\r\n000000算量::站房部分::9.9::候车大厅::其他墙面::<N1>干挂花岗岩内墙面:平方米\r\n000000算量::站房部分::9.9::候车大厅::其他墙面::GRG柱面:平方米\r\n000000算量::站房部分::9.9::候车大厅::其他墙面::风柱铝板(有改动):平方米\r\n000000算量::站房部分::9.9::候车大厅::玻璃墙面\r\n000000算量::站房部分::9.9::候车大厅::玻璃墙面::<Q8>4.5m 8+1.52PVB+8:平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶\r\n000000算量::站房部分::9.9::候车大厅::吊顶::<C2>密拼铝板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶::<C2>密拼木色纯色铝板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶::<C3>1.4m间距550x50格栅铝板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶::<C2>周边收边铝单板吊顶:平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶::<C2>夹层铝单板吊顶:平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶::<C2>门斗铝单板吊顶:平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶翻边\r\n000000算量::站房部分::9.9::候车大厅::吊顶翻边::<C2>立面柱侧面密拼铝板:平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶翻边::<C2>夹层檐口铝单板(有改动):平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶翻边::<C2>夹层檐口蓝色纯色铝单板(有改动):平方米\r\n000000算量::站房部分::9.9::候车大厅::吊顶翻边::<C2>天窗侧面铝单板:平方米\r\n000000算量::站房部分::9.9::候车大厅::地面\r\n000000算量::站房部分::9.9::候车大厅::地面::<F1A>花岗岩楼面:平方米\r\n000000算量::站房部分::9.9::候车大厅::其他工程\r\n000000算量::站房部分::9.9::候车大厅::其他工程::<Q8>2.2m 8+1.52PVB+8:米\r\n000000算量::站房部分::9.9::候车大厅::其他工程::<Q8>1.4m 8+1.52PVB+8:米\r\n000000算量::站房部分::9.9::候车大厅::其他工程::单向门:个\r\n000000算量::站房部分::9.9::候车大厅::其他工程::定制服务台:个\r\n000000算量::站房部分::9.9::候车大厅::其他工程::6.4m*3m 疏散感应门:个\r\n000000算量::站房部分::9.9::候车大厅::其他工程::幕墙底部不锈钢踢脚:米\r\n000000算量::站房部分::9.9::候车大厅::其他工程::150mm高踢脚:米\r\n000000算量::站房部分::9.9::候车大厅::其他工程::玻璃下防撞踢脚:米\r\n000000算量::站房部分::9.9::上商业夹层楼梯间\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::其他墙面\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::其他墙面::<N2A>穿孔铝板(吸音板)内墙面:平方米\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::吊顶\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::吊顶::<C2>1.5mm铝条板:平方米\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::地面\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::地面::<F1B>花岗岩楼面:平方米\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::栏杆\r\n000000算量::站房部分::9.9::上商业夹层楼梯间::栏杆::<Q8>1.1m 8+1.52PVB+8玻璃栏板:米\r\n000000算量::站房部分::9.9::售票厅\r\n000000算量::站房部分::9.9::售票厅::其他墙面\r\n000000算量::站房部分::9.9::售票厅::其他墙面::<N1>干挂花岗岩内墙面:平方米\r\n000000算量::站房部分::9.9::售票厅::吊顶\r\n000000算量::站房部分::9.9::售票厅::吊顶::<C2>纯色木色铝单板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::售票厅::吊顶::<C9>石膏造型板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::售票厅::地面\r\n000000算量::站房部分::9.9::售票厅::地面::<F1A>花岗岩楼面:平方米\r\n000000算量::站房部分::9.9::售票厅::其他工程\r\n000000算量::站房部分::9.9::售票厅::其他工程::服务台:个\r\n000000算量::站房部分::9.9::行政贵宾\r\n000000算量::站房部分::9.9::行政贵宾::其他墙面\r\n000000算量::站房部分::9.9::行政贵宾::其他墙面::<N4>干挂大理石墙面:平方米\r\n000000算量::站房部分::9.9::行政贵宾::吊顶\r\n000000算量::站房部分::9.9::行政贵宾::吊顶::<C9>造型石膏板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::行政贵宾::地面\r\n000000算量::站房部分::9.9::行政贵宾::地面::<F7A>大理石楼面:平方米\r\n000000算量::站房部分::9.9::行政贵宾::其他工程\r\n000000算量::站房部分::9.9::行政贵宾::其他工程::150mm高踢脚:米\r\n000000算量::站房部分::9.9::行政贵宾::其他工程::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::站房部分::9.9::行政贵宾::其他工程::沙发(有改动):个\r\n000000算量::站房部分::9.9::贵宾候车厅\r\n000000算量::站房部分::9.9::贵宾候车厅::其他墙面\r\n000000算量::站房部分::9.9::贵宾候车厅::其他墙面::<N4>花岗岩墙面(有改动):平方米\r\n000000算量::站房部分::9.9::贵宾候车厅::吊顶\r\n000000算量::站房部分::9.9::贵宾候车厅::吊顶::<C9>1mm铝条板吊顶(有改动):平方米\r\n000000算量::站房部分::9.9::贵宾候车厅::地面\r\n000000算量::站房部分::9.9::贵宾候车厅::地面::<F7A>花岗岩地面(有改动):平方米\r\n000000算量::站房部分::9.9::贵宾候车厅::其他工程\r\n000000算量::站房部分::9.9::贵宾候车厅::其他工程::贵宾150mm高踢脚:米\r\n000000算量::站房部分::9.9::贵宾候车厅::其他工程::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::站房部分::9.9::贵宾候车厅::其他工程::整套沙发家具(有改动):个\r\n000000算量::站房部分::9.9::进站检票口\r\n000000算量::站房部分::9.9::进站检票口::其他墙面\r\n000000算量::站房部分::9.9::进站检票口::其他墙面::<N2>干挂铝板侧墙(有改动):平方米\r\n000000算量::站房部分::9.9::进站检票口::吊顶\r\n000000算量::站房部分::9.9::进站检票口::吊顶::<C2>铝单板吊顶:平方米\r\n000000算量::站房部分::9.9::进站检票口::地面\r\n000000算量::站房部分::9.9::进站检票口::地面::<F1A>花岗岩楼面:平方米\r\n000000算量::站房部分::9.9::进站检票口::其他工程\r\n000000算量::站房部分::9.9::进站检票口::其他工程::<Q8>2.2m 8+1.52PVB+8:米\r\n000000算量::站房部分::9.9::进站检票口::其他工程::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::站房部分::9.9::进站楼梯\r\n000000算量::站房部分::9.9::进站楼梯::地面\r\n000000算量::站房部分::9.9::进站楼梯::地面::<Q3>进出站楼梯踢面踏面:平方米\r\n000000算量::站房部分::9.9::进站楼梯::栏杆\r\n000000算量::站房部分::9.9::进站楼梯::栏杆::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::站房部分::17.4\r\n000000算量::站房部分::17.4::商业夹层\r\n000000算量::站房部分::17.4::商业夹层::其他墙面\r\n000000算量::站房部分::17.4::商业夹层::其他墙面::<N1>干挂花岗岩内墙面:平方米\r\n000000算量::站房部分::17.4::商业夹层::玻璃墙面\r\n000000算量::站房部分::17.4::商业夹层::玻璃墙面::<Q8>4.5m 8+1.52PVB+8:平方米\r\n000000算量::站房部分::17.4::商业夹层::吊顶\r\n000000算量::站房部分::17.4::商业夹层::吊顶::<C2>铝单板吊顶:平方米\r\n000000算量::站房部分::17.4::商业夹层::吊顶翻边\r\n000000算量::站房部分::17.4::商业夹层::吊顶翻边::<C2>铝单板吊顶:平方米\r\n000000算量::站房部分::17.4::商业夹层::地面\r\n000000算量::站房部分::17.4::商业夹层::地面::<F1A>花岗岩楼面:平方米\r\n000000算量::站房部分::17.4::商业夹层::地面::钢结构顶板2小时防火涂料:平方米\r\n000000算量::站房部分::17.4::商业夹层::栏杆\r\n000000算量::站房部分::17.4::商业夹层::栏杆::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::站房部分::17.4::商业夹层::栏杆::<Q8>1.3m 8+1.52PVB+8:米\r\n000000算量::站房部分::17.4::商业夹层::其他工程\r\n000000算量::站房部分::17.4::商业夹层::其他工程::玻璃下防撞踢脚:米\r\n000000算量::出站地道\r\n000000算量::出站地道::-9.0\r\n000000算量::出站地道::-9.0::出站地道\r\n000000算量::出站地道::-9.0::出站地道::其他墙面\r\n000000算量::出站地道::-9.0::出站地道::其他墙面::<N1>干挂花岗岩内墙面:平方米\r\n000000算量::出站地道::-9.0::出站地道::吊顶\r\n000000算量::出站地道::-9.0::出站地道::吊顶::<C2>密拼铝板吊顶:平方米\r\n000000算量::出站地道::-9.0::出站地道::地面\r\n000000算量::出站地道::-9.0::出站地道::地面::<F1C>花岗岩楼面三:平方米\r\n000000算量::出站地道::-9.0::出站地道::地面::300宽排水沟:米\r\n000000算量::出站地道::-9.0::出站楼梯\r\n000000算量::出站地道::-9.0::出站楼梯::其他墙面\r\n000000算量::出站地道::-9.0::出站楼梯::其他墙面::<N1>干挂花岗岩内墙面:平方米\r\n000000算量::出站地道::-9.0::出站楼梯::其他墙面::<Q3>出站楼梯踢面踏面:平方米\r\n000000算量::出站地道::-9.0::出站楼梯::吊顶\r\n000000算量::出站地道::-9.0::出站楼梯::吊顶::<C2>密拼铝板吊顶:平方米\r\n000000算量::出站地道::-9.0::出站楼梯::地面\r\n000000算量::出站地道::-9.0::出站楼梯::地面::<F1C>花岗岩楼面三:平方米\r\n000000算量::出站地道::-9.0::出站楼梯::栏杆\r\n000000算量::出站地道::-9.0::出站楼梯::栏杆::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::架空换乘区\r\n000000算量::架空换乘区::-9.0\r\n000000算量::架空换乘区::-9.0::换乘大厅\r\n000000算量::架空换乘区::-9.0::换乘大厅::其他墙面\r\n000000算量::架空换乘区::-9.0::换乘大厅::其他墙面::<W1B>干挂铝板柱面:平方米\r\n000000算量::架空换乘区::-9.0::换乘大厅::吊顶\r\n000000算量::架空换乘区::-9.0::换乘大厅::吊顶::<C1>穿孔铝板 造型铝板吊顶:平方米\r\n000000算量::架空换乘区::-9.0::换乘大厅::地面\r\n000000算量::架空换乘区::-9.0::换乘大厅::地面::<F1C>花岗岩楼面三:平方米\r\n000000算量::架空换乘区::-9.0::旅服中庭\r\n000000算量::架空换乘区::-9.0::旅服中庭::其他墙面\r\n000000算量::架空换乘区::-9.0::旅服中庭::其他墙面::<W1B>干挂铝板柱面:平方米\r\n000000算量::架空换乘区::-9.0::旅服中庭::吊顶\r\n000000算量::架空换乘区::-9.0::旅服中庭::吊顶::<C1>穿孔铝板 造型铝板吊顶:平方米\r\n000000算量::架空换乘区::-9.0::旅服中庭::地面\r\n000000算量::架空换乘区::-9.0::旅服中庭::地面::<F1C>花岗岩楼面三:平方米\r\n000000算量::架空换乘区::-9.0::旅服中庭::其他工程\r\n000000算量::架空换乘区::-9.0::旅服中庭::其他工程::1.5m挡烟垂壁:米\r\n000000算量::架空换乘区::-9.0::旅服中庭::其他工程::玻璃下防撞踢脚:米\r\n000000算量::室外平台\r\n000000算量::室外平台::0.0\r\n000000算量::室外平台::0.0::室外活动平台\r\n000000算量::室外平台::0.0::室外活动平台::其他墙面\r\n000000算量::室外平台::0.0::室外活动平台::其他墙面::<W2>干挂铝板柱面:平方米\r\n000000算量::室外平台::0.0::室外活动平台::地面\r\n000000算量::室外平台::0.0::室外活动平台::地面::<G2B>落客平台花岗岩楼面:平方米\r\n000000算量::室外平台::0.0::室外活动平台::地面::<G2B>贵宾车道沥青楼面:平方米\r\n000000算量::室外平台::0.0::室外活动平台::地面::<Q3>进出站楼梯踢面踏面:平方米\r\n000000算量::室外平台::0.0::室外活动平台::栏杆\r\n000000算量::室外平台::0.0::室外活动平台::栏杆::<Q8>1.3m 8+1.52PVB+8:米\r\n000000算量::室外平台::0.0::室外活动平台::栏杆::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::室外平台::0.0::室外活动平台::其他工程\r\n000000算量::室外平台::0.0::室外活动平台::其他工程::500宽排水沟:米\r\n000000算量::室外平台::0.0::室外活动平台::其他工程::3.0m 8+1.52PVB+8:米\r\n000000算量::室外平台::0.0::室外活动平台::其他工程::5.4m宽3m高不锈钢管理门:个\r\n000000算量::室外平台::0.0::室外活动平台::其他工程::1500间距800直径防撞墩:个\r\n000000算量::室外平台::0.0::室外活动平台::其他工程::150mm高踢脚:米\r\n000000算量::室外平台::0.0::站台\r\n000000算量::室外平台::0.0::站台::站台地面\r\n000000算量::室外平台::0.0::站台::站台地面::基本站台铺面总面积:平方米\r\n000000算量::室外平台::0.0::站台::站台地面::中间站台铺面总面积:平方米\r\n000000算量::室外平台::0.0::站台::站台地面::<G3>30mm中间站台铺面:平方米\r\n000000算量::室外平台::0.0::站台::站台地面::<G3>50mm基本站台铺面:平方米\r\n000000算量::室外平台::0.0::站台::站台地面::600mm宽50mm厚机刨石:平方米\r\n000000算量::室外平台::0.0::站台::站台地面::600mm宽盲道:平方米\r\n000000算量::室外平台::0.0::站台::站台地面::100mm宽汉白玉安全线:平方米\r\n000000算量::室外平台::0.0::站台::其他墙面\r\n000000算量::室外平台::0.0::站台::其他墙面::雨棚柱装饰涂料(有改动):平方米\r\n000000算量::室外平台::0.0::站台::吊顶\r\n000000算量::室外平台::0.0::站台::吊顶::雨棚下装饰涂料(有改动):平方米\r\n000000算量::室外平台::0.0::站台::吊顶::<R2>混凝土屋面(有改动):平方米\r\n000000算量::室外平台::0.0::站台::其他工程\r\n000000算量::室外平台::0.0::站台::其他工程::端部爬梯(有改动):米\r\n000000算量::室外平台::0.0::站台::其他工程::1.4m高端部不锈钢栏杆:米\r\n000000算量::室外平台::9.9\r\n000000算量::室外平台::9.9::落客平台\r\n000000算量::室外平台::9.9::落客平台::其他墙面\r\n000000算量::室外平台::9.9::落客平台::其他墙面::清水混凝土柱:平方米\r\n000000算量::室外平台::9.9::落客平台::地面\r\n000000算量::室外平台::9.9::落客平台::地面::<Q3>进出站楼梯踢面踏面:平方米\r\n000000算量::室外平台::9.9::落客平台::地面::<G2B>落客平台花岗岩楼面(有改动):平方米\r\n000000算量::室外平台::9.9::落客平台::栏杆\r\n000000算量::室外平台::9.9::落客平台::栏杆::<Q8>1.1m 8+1.52PVB+8:米\r\n000000算量::室外平台::9.9::落客平台::其他工程\r\n000000算量::室外平台::9.9::落客平台::其他工程::500宽排水沟(有改动):米\r\n000000算量::室外平台::9.9::落客平台::其他工程::1500间距800直径防撞墩:个\r\n000000算量::室外平台::9.9::落客平台::其他工程::平台与高架连接处侧缘石:米\r\n000000算量::室外平台::9.9::员工停车场\r\n000000算量::室外平台::9.9::员工停车场::其他工程\r\n000000算量::室外平台::9.9::员工停车场::其他工程::5.4m宽3m高不锈钢管理门(有改动):个\r\n000000算量::室外平台::9.9::员工停车场::其他工程::<Q8>3.0m 8+1.52PVB+8(有改动):米\r\n000000算量::室外平台::9.9::员工停车场::其他工程::500宽排水沟(有改动):米\r\n000000算量::室外平台::9.9::员工停车场::地面\r\n000000算量::室外平台::9.9::员工停车场::地面::<G2B>落客平台花岗岩楼面(有改动):平方米\r\n000000算量::室外平台::9.9::临时停车场\r\n000000算量::室外平台::9.9::临时停车场::其他工程\r\n000000算量::室外平台::9.9::临时停车场::其他工程::5.4m宽3m高不锈钢管理门(有改动):个\r\n000000算量::室外平台::9.9::临时停车场::其他工程::<Q8>3.0m 8+1.52PVB+8(有改动):米\r\n000000算量::室外平台::9.9::临时停车场::其他工程::500宽排水沟(有改动):米\r\n000000算量::室外平台::9.9::临时停车场::地面\r\n000000算量::室外平台::9.9::临时停车场::地面::<G2B>落客平台花岗岩楼面(有改动):平方米\r\n000000算量::高架层底面吊顶\r\n000000算量::高架层底面吊顶::0.0\r\n000000算量::高架层底面吊顶::0.0::站台\r\n000000算量::高架层底面吊顶::0.0::站台::吊顶\r\n000000算量::高架层底面吊顶::0.0::站台::吊顶::高架下方涂料水平投影面:平方米\r\n000000算量::铁路停车场\r\n000000算量::铁路停车场::0.0\r\n000000算量::铁路停车场::0.0::地面铁路停车场\r\n000000算量::铁路停车场::0.0::地面铁路停车场::地面\r\n000000算量::铁路停车场::0.0::地面铁路停车场::地面::<G4A>车库混凝土地面:平方米\r\n000000算量::铁路停车场::0.0::地面铁路停车场::地面::绿化:平方米\r\n000000算量::铁路停车场::0.0::地面铁路停车场::其他工程\r\n000000算量::铁路停车场::0.0::地面铁路停车场::其他工程::300宽排水沟:米\r\n000000算量::铁路停车场::0.0::地面铁路停车场::其他工程::2.2m不锈钢栏杆:米\r\n000000算量::铁路停车场::0.0::地面铁路停车场::其他工程::进出口管理岗亭:个\r\n000000算量::铁路停车场::-9.0\r\n000000算量::铁路停车场::-9.0::地下铁路停车场\r\n000000算量::铁路停车场::-9.0::地下铁路停车场::其他工程\r\n000000算量::铁路停车场::-9.0::地下铁路停车场::其他工程::300宽排水沟:米\r\n000000算量::铁路停车场::-9.0::地下铁路停车场::地面\r\n000000算量::铁路停车场::-9.0::地下铁路停车场::地面::<G4>车库混凝土地面:平方米\r\ncad\r\ncad::temp\r\ncad::PUB_TEXT\r\ncad::DL-缘石线\r\ncad::家具\r\ncad::WALL\r\ncad::JZ-5-家具 布置\r\ncad::STAIR\r\ncad::WINDOW\r\ncad::SPACE\r\ncad::DIM_SYMB\r\ncad::DOOR_FIRE\r\ncad::DOOR_FIRE_TEXT\r\ncad::看线1\r\ncad::BZ\r\ncad::配景\r\ncad::9-F-Q轴立面图$0$JZ-5-装修线\r\ncad::CURTWALL\r\ncad::站场平面_曲线头\r\ncad::站台\r\ncad::0-大理北站工区（施工图）\r\ncad::【站场设计】更改\r\ncad::√贯通正线D1K-考虑大攀160-坡度折减-大丽线控制\r\ncad::【施工图】改建大理北站至大理站联络线\r\ncad::【施工图】改建大丽线\r\ncad::【ZC-道岔编号】\r\ncad::0-站场平面\r\ncad::ZC-大理北客站涵洞插旗\r\ncad::JZ-2-看线\r\ncad::SN-铺装\r\ncad::000-cars\r\ncad::SPACE_HATCH\r\ncad::SZY$ROOF$01LINE$1\r\ncad::PUB_HATCH\r\ncad::Make2D$可见线$曲线$路缘石\r\ncad::TEXT\r\ncad::DOTE\r\ncad::COLUMN\r\ncad::JZ-5-装修线\r\ncad::temp(不打印)\r\ncad::PUB_DIM\r\ncad::DIM_ELEV\r\ncad::JZ-0-设备\r\ncad::REDLINE\r\ncad::C\r\ncad::A-FLOR-FURN\r\ncad::JZ-5-扶手\r\ncad::DIM_LEAD\r\ncad::z-配景\r\ncad::0-股道顺坡\r\ncad::JZ-5-绿化\r\ncad::建筑-栏杆、扶手\r\ncad::设备-设备\r\ncad::JZ-客运设备\r\ncad::1-楼梯\r\ncad::A-FLOR-电梯\r\ncad::AZ-楼扶梯洞口\r\ncad::PROMPT\r\ncad::1-lyd\r\ncad::HV-提建筑\r\ncad::HV-K1-SB\r\ncad::HV-F1-SF\r\ncad::z-结构板控制线（不打印）\r\ncad::car\r\ncad::Make2D$可见线$曲线$楼-站前\r\ncad::Make2D$可见线$曲线$景观-站前\r\ncad::中心线\r\ncad::17.40\r\ncad::17.40::0\r\ncad::17.40::WALL\r\ncad::17.40::家具\r\ncad::17.40::看线1\r\ncad::17.40::PUB_TEXT\r\ncad::17.40::DIM_SYMB\r\ncad::17.40::STAIR\r\ncad::17.40::DOOR_FIRE\r\ncad::17.40::DOOR_FIRE_TEXT\r\ncad::17.40::JZ-2-看线\r\ncad::17.40::JZ-5-家具 布置\r\ncad::17.40::SPACE\r\ncad::17.40::WINDOW\r\ncad::17.40::CURTWALL\r\ncad::17.40::PUB_HATCH\r\ncad::17.40::DIM_LEAD\r\ncad::17.40::PUB_DIM\r\ncad::17.40::DOTE\r\ncad::17.40::SURFACE\r\ncad::17.40::COLUMN\r\ncad::17.40::DIM_ELEV\r\ncad::17.40::temp(不打印)\r\ncad::17.40::temp\r\ncad::0\r\nZNN\r\nZNN::雨棚柱\r\n";
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