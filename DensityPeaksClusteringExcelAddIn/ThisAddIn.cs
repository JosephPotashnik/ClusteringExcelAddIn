using System;
using DensityPeaksClustering;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Worksheet = Microsoft.Office.Tools.Excel.Worksheet;

namespace DensityPeaksClusteringExcelAddIn
{
    public partial class ThisAddIn
    {
        private Range _target;

        private readonly int[] _colors =
        {
            (int) XlRgbColor.rgbBlack,
            (int) XlRgbColor.rgbDarkRed,
            (int) XlRgbColor.rgbGreen,
            (int) XlRgbColor.rgbBlueViolet,
            (int) XlRgbColor.rgbYellow,
            (int) XlRgbColor.rgbAqua,
            (int) XlRgbColor.rgbOrangeRed,
            (int) XlRgbColor.rgbSandyBrown,
            (int) XlRgbColor.rgbLightSalmon,
            (int) XlRgbColor.rgbLightGray,
            (int) XlRgbColor.rgbLightPink,
            (int) XlRgbColor.rgbMidnightBlue,
            (int) XlRgbColor.rgbLavender,
            (int) XlRgbColor.rgbMediumSpringGreen,
            (int) XlRgbColor.rgbViolet
        };

        private void DensityPeaksClusteringMenuItemClick(Office.CommandBarButton ctrl, ref bool cancelDefault)
        {
            //check whether we have a labels row
            if (!double.TryParse(_target.Cells[1][1].Value.ToString(), out double num))
            {
                _target = _target.Offset[1];
                _target =  _target.Resize[_target.Rows.Count - 1];
            }


            //create matrix
            var numOfSamples = _target.Rows.Count;
            var numOfDimensions = _target.Columns.Count;
            var m = CreateMatrixFromRange(numOfDimensions, numOfSamples);

            //cluster
            var clusters = DensityPeaksClusteringAlgorithms.MultiManifold(m);

            //create chart with color coded clusters
            var worksheet = CreateXYScatterChartWithClusters(clusters);

            worksheet.get_Range("A1", "A1").Select();
        }

        private Worksheet CreateXYScatterChartWithClusters(int[] clusters)
        {
            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            if (worksheet.Controls.Contains("clustering"))
                worksheet.Controls.Remove("clustering");

            var chart = worksheet.Controls.AddChart(150, 10, 450, 400, "clustering");

            chart.ChartType = XlChartType.xlXYScatter;
            chart.HasLegend = false;

            Axis vertAxis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            vertAxis.HasMajorGridlines = false; 
            vertAxis.MaximumScaleIsAuto = true;
            vertAxis.MinimumScaleIsAuto = true;

            chart.SetSourceData(_target);
            var series1 = (Series) chart.SeriesCollection(1);

            for (var i = 0; i < clusters.Length; i++)
            {
                var p = (Point) series1.Points(i + 1 );
                p.MarkerBackgroundColor = _colors[clusters[i] % _colors.Length];
            }

            return worksheet;
        }

        private float[][] CreateMatrixFromRange(int numOfDimensions, int numOfSamples)
        {
            var m = new float[numOfSamples][];
            for (var i = 0; i < numOfSamples; i++)
            {
                m[i] = new float[numOfDimensions];
                for (var j = 0; j < numOfDimensions; j++)
                    m[i][j] = (float) _target.Cells[ j + 1][i + 1].Value;
            }

            return m;
        }

        private void AddDensityPeaksClusteringMenuItem()
        {
            var menuItem = Office.MsoControlType.msoControlButton;
            var densityPeaksClusteringMenuItem =
                (Office.CommandBarButton) GetCellContextMenu().Controls.Add(menuItem, missing, missing, 1, true);

            densityPeaksClusteringMenuItem.Style = Office.MsoButtonStyle.msoButtonCaption;
            densityPeaksClusteringMenuItem.Caption = "Density Peaks Clustering";
            densityPeaksClusteringMenuItem.Click += DensityPeaksClusteringMenuItemClick;
        }

        private void Application_SheetBeforeRightClick(object sh, Range target, ref bool cancel)
        {
            ResetCellMenu(); // reset the cell context menu back to the default

            if (target.Cells.Columns.Count == 2)
            {
                _target = target;
                AddDensityPeaksClusteringMenuItem();
            }
        }

        private void ResetCellMenu()
        {
            GetCellContextMenu().Reset();
        }

        private Office.CommandBar GetCellContextMenu()
        {
            return Application.CommandBars["Cell"];
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ResetCellMenu(); // reset the cell context menu back to the default

            // Call this function is the user right clicks on a cell
            Application.SheetBeforeRightClick += Application_SheetBeforeRightClick;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}