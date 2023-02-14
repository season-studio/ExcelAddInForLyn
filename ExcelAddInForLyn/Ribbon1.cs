using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

// TODO:   按照以下步骤启用功能区(XML)项:

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace ExcelAddInForLyn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddInForLyn.Ribbon1.xml");
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region 自定义功能

        public void DoPickSchedule(bool _hortizontal)
        {
            const int MaxContinuedEmptyCount = 100;

            try
            {
                if (Globals.ThisAddIn.Application.Selection is Range range)
                {
                    if (range.Areas.Count > 1)
                    {
                        MessageBox.Show("不支持多个不连续区间的选择", Globals.ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var startAddress = range[1, 1].Address; ;
                    // var endAddress = range.SpecialCells(XlCellType.xlCellTypeLastCell).Address;
                    // range = (Globals.ThisAddIn.Application.ActiveSheet as Worksheet).Range[$"{startAddress}:{endAddress}"];

                    int rowsCount = range.Rows.Count, colsCount = range.Columns.Count;
                    if ((rowsCount < 2) || (colsCount < 2))
                    {
                        MessageBox.Show("没有选择足够的数据。", Globals.ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    int seriesIdx = 0, dataIdx = 0;
                    Func<string> fnGetData;
                    Func<bool> fnStepAndCheckSeries;
                    Func<bool> fnStepAndCheckData;
                    Func<string> fnGetScheduleDate;
                    Func<string> fnGetDataAddress;
                    if (_hortizontal)
                    {
                        fnGetData = () => range.Cells[seriesIdx, dataIdx].Text;
                        fnStepAndCheckSeries = () => (++seriesIdx) <= rowsCount;
                        fnStepAndCheckData = () => (++dataIdx) <= colsCount;
                        fnGetScheduleDate = () => range.Cells[1, dataIdx].Text;
                        fnGetDataAddress = () => (range.Cells[seriesIdx, dataIdx] as Range).Address;
                    }
                    else
                    {
                        fnGetData = () => range.Cells[dataIdx, seriesIdx].Text;
                        fnStepAndCheckSeries = () => (++seriesIdx) <= colsCount;
                        fnStepAndCheckData = () => (++dataIdx) <= rowsCount;
                        fnGetScheduleDate = () => range.Cells[dataIdx, 1].Text;
                        fnGetDataAddress = () => (range.Cells[dataIdx, seriesIdx] as Range).Address;
                    }

                    List<string[]> allSchedules = new List<string[]>();
                    List<string> itemSchedules = new List<string>();
                    var endAddress = startAddress;
                    int emptyCount = 0;
                    seriesIdx = 2;
                    do
                    {
                        dataIdx = 1;
                        string seriesName = fnGetData();
                        if (string.IsNullOrWhiteSpace(seriesName))
                        {
                            if (++emptyCount >= MaxContinuedEmptyCount)
                            {
                                break;
                            }
                        }
                        else
                        {
                            emptyCount = 0;
                            itemSchedules.Clear();
                            while (fnStepAndCheckData())
                            {
                                string val = fnGetData();
                                endAddress = fnGetDataAddress();
                                if (!string.IsNullOrWhiteSpace(val))
                                {
                                    itemSchedules.Add(string.Join("\x1B", fnGetScheduleDate(), val));
                                }
                            }
                            if (itemSchedules.Count > 0)
                            {
                                allSchedules.Add(itemSchedules.Prepend(seriesName).ToArray());
                            }
                        }
                    } while (fnStepAndCheckSeries());

                    var newSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add() as Worksheet;
                    if (null != newSheet)
                    {
                        var newRange = newSheet.Range[$"A1:C1"];
                        newRange[1, 1] = $"数据源范围{startAddress}:{endAddress}";
                        newRange[1, 2] = "排期时间";
                        newRange[1, 3] = "数量";
                        newRange = newRange.Offset[1, 0];
                        foreach (var seriesRecords in allSchedules)
                        {
                            if (seriesRecords.Length > 1)
                            {
                                var seriesName = seriesRecords[0];
                                for (int vIdx = 1; vIdx < seriesRecords.Length; vIdx++)
                                {
                                    newRange[1, 1] = seriesName;
                                    var values = seriesRecords[vIdx].Split('\x1B');
                                    newRange[1, 2] = values[0];
                                    if (!string.IsNullOrWhiteSpace(values[1]))
                                    {
                                        newRange[1, 3] = values[1];
                                    }
                                    newRange = newRange.Offset[1, 0];
                                }
                            }
                        }

                        newSheet.Name = $"抽取结果({DateTime.Now.ToString("yy-MM-dd HH.mm.ss.fff")})";
                        newSheet.Activate();
                    }

                    MessageBox.Show($"数据处理完成。{Environment.NewLine}有效的源数据区间是{startAddress}:{endAddress}。{Environment.NewLine}建议检查一下有效数据是否与选择匹配。", Globals.ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // MessageBox.Show(string.Join(Environment.NewLine, allSchedules.Select(e => $"[{string.Join(", ", e)}]")));
                }
                else
                {
                    MessageBox.Show("需要选择源数据区间", Globals.ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception err)
            {
                Trace.TraceError(err.ToString());
                MessageBox.Show($"出现错误。{Environment.NewLine}{err.Message}", Globals.ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnPickProductScheduleHorizontal(IRibbonControl _)
        {
            DoPickSchedule(true);
        }

        public void OnPickProductScheduleVeritcal(IRibbonControl _)
        {
            DoPickSchedule(false);
        }
        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
