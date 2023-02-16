using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelAddInForLyn
{
    internal class Spotlight
    {
        private static readonly string ConfigColorName = "SpotlightColor";

        public Spotlight()
        {
            lastSheet = null;
            enableSpotlight = false;
            color = Color.FromArgb((int)Configuration.Get(ConfigColorName, (int)0xFFCCAA));
            Globals.ThisAddIn.Application.SheetSelectionChange += OnSheetSelectionChange;
        }

        private Worksheet lastSheet;

        private bool enableSpotlight;

        private Color color;

        public Color Color
        {
            get => color;
            set
            {
                color = value;
                Configuration.Set(ConfigColorName, color.ToArgb());
            }
        }

        public bool Enable
        {
            get => enableSpotlight;
            set
            {
                if (enableSpotlight ^ value)
                {
                    enableSpotlight = value;
                    if (!value && (null != lastSheet))
                    {
                        try
                        {
                            lastSheet.Cells.Interior.ColorIndex = 0;
                        }
                        catch (Exception err)
                        {
                            Trace.TraceWarning(err.ToString());
                        }
                        finally
                        {
                            Interlocked.Exchange(ref lastSheet, null);
                        }
                    }
                }
            }
        }

        private void OnSheetSelectionChange(object Sh, Range Target)
        {
            if ((Sh is Worksheet sheet) && enableSpotlight)
            {
                if (lastSheet != sheet)
                {
                    try
                    {
                        if (null != lastSheet)
                        {
                            lastSheet.Cells.Interior.ColorIndex = 0;
                        }
                    }
                    catch (Exception err)
                    {
                        Trace.TraceWarning(err.ToString());
                    }
                    finally
                    {
                        Interlocked.Exchange(ref lastSheet, sheet);
                    }
                }
                try
                {
                    sheet.Cells.Interior.ColorIndex = 0;
                    var addrs = Target.Address[true, false].Split('$');
                    var range = Globals.ThisAddIn.Application.Union(sheet.Range[$"{addrs[0]}:{addrs[0]}"], sheet.Range[$"{addrs[1]}:{addrs[1]}"]);
                    range.Interior.Color = color;
                }
                catch (Exception err)
                {
                    Trace.TraceWarning(err.ToString());
                }
            }
        }
    }
}
