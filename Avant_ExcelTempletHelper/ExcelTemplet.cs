using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Avant_ExcelTempletHelper
{
    public class ExcelTemplet
    {
        WorkbookDesigner WD = new WorkbookDesigner();

        string Path;
        string TempletPath;
        public void PutImage(string Path, int LeftRow=0,int LeftCoulmn=0,int RightRow=1,int RightCoulmn=1)
        {
            Workbook workbook = WD.Workbook;
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells[LeftRow,LeftCoulmn].PutValue("Image Hyperlink");
            int index = worksheet.Pictures.Add(LeftRow, LeftCoulmn, RightRow, RightCoulmn, Path);
            Aspose.Cells.Drawing.Picture pic = worksheet.Pictures[index];
            pic.Placement = PlacementType.FreeFloating;
            Aspose.Cells.Hyperlink hlink = pic.AddHyperlink(Path);
            hlink.ScreenTip = "Click this Picture to go to view Source File";
        }
        public void PutImage(Stream ImageSteam,int LeftRow = 0, int LeftCoulmn = 0, int RightRow = 1, int RightCoulmn = 1)
        {
            Workbook workbook = WD.Workbook;
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells[LeftRow, LeftCoulmn].PutValue("Image Hyperlink");
            int index = worksheet.Pictures.Add(LeftRow, LeftCoulmn, RightRow, RightCoulmn, ImageSteam);
            Aspose.Cells.Drawing.Picture pic = worksheet.Pictures[index];
            pic.Placement = PlacementType.FreeFloating;
            Aspose.Cells.Hyperlink hlink = pic.AddHyperlink(Path);
            hlink.ScreenTip = "Click this Picture to go to view Source File";
        }
        /// <summary>
        /// 保存文件的类型
        /// </summary>
        public SaveFormat _saveFormat;
        /// <summary>
        /// 保存Excel文件的路径
        /// </summary>
        public string SavePath
        {
            set { this.Path = value; }
            get { return this.Path; }
        }
        /// <summary>
        /// 源EXCEL模板路径
        /// </summary>
        public string TempletFilePath
        {
            set { this.TempletPath = value; }
            get { return this.TempletPath; }
        }
        /// <summary>
        /// 初始化模板信息
        /// </summary>
        /// <param name="FileName">需要保存的EXCLE文件名</param>
        /// <param name="Templet">EXCLE模板文件</param>
        public ExcelTemplet(string FileName, string Templet)
        {
            Path = FileName;
            TempletPath = Templet;
            WD.Workbook = new Workbook(TempletPath);  
        }
        /// <summary>
        /// 增加数据源到模板中
        /// </summary>
        /// <param name="DT">源数据</param>
        /// <param name="TableName">变量名。Ex：&=[TableName].[DT中的列名]</param>
        public void SetDataTableToExcelTemplet(DataTable DT, string TableName)
        {
            DT.TableName = TableName;
            try { WD.SetDataSource(DT); } catch { }

        }
        /// <summary>
        /// 增加单个变量到模板
        /// </summary>
        /// <param name="Variable">变量名。Ex：&=$变量名</param>
        /// <param name="Value">值</param>
        public void SetStringToExcelTemplet(string Variable, string Value)
        {
            WD.SetDataSource(Variable, Value);

        }
        /// <summary>
        /// 批量增加数据集合到模板
        /// </summary>
        /// <param name="Values">List<Tuple<&=$变量名,数据>> </param>
        public void SetStringToExcelTemplet(List<Tuple<string, object>> Values)
        {
            foreach (Tuple<string, object> Temp in Values)
            {

                WD.SetDataSource(Temp.Item1, Temp.Item2);

            }
        }
        /// <summary>
        /// 保存文件
        /// </summary>
        public void Save()
        {
            WD.Process();
            WD.Workbook.Save(Path);
          
        }
        /// <summary>
        /// 保存文件，避免独占方式打开后再次保存抛出的异常
        /// </summary>
        public string SafeSave()
        { 
            try
            {
                Save();
                return "OK";
            }
            catch (Exception ex)
            { 
                return ex.Message.ToString();
            }
        }
    }
}
