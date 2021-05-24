using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Highlighter
{
    delegate void myDelegate(Control ctl, object i, object max);
    public partial class Form1 : Form
    {
        int max = 10000;
        int startRowData = 1;
        List<string> paths = new List<string>();
        List<DataTable> DTs = new List<DataTable>();
        public Form1()
        {
            InitializeComponent();
            this.ActivateMdiChild(this);
            txt_before.ResetText();
            txt_current.ResetText();
        }

        //전월 엑셀파일 선택
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txt_before.Text = dialog.FileName;
            }

        }

        //당월 엑셀파일 선택
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txt_current.Text = dialog.FileName;
            }
        }

        //엑셀 작업 시작!
        private void button3_Click(object sender, EventArgs e)
        {
            if (txt_before.Text.Equals("") || txt_current.Text.Equals(""))
            {
                MessageBox.Show("엑셀파일을 선택 해 주세요.");
                return;
            }
            ThreadStart working = new ThreadStart(doWork);
            Thread working_thread = new Thread(working);
            working_thread.Start();
        }

        void initProgress()
        {
            pb_loading.Maximum = max;
            pb_loading.Value = 0;
        }

        void doWork()
        {
            updateProgress(this, 1, max);
            //엑셀오픈
            Excel.Application xlapp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws1 = null;

            paths.Add(txt_before.Text.Trim());
            paths.Add(txt_current.Text.Trim());

            DataTable dtBeforeMonth = new DataTable();
            DataTable dtCurrentMonth = new DataTable();
            DTs.Add(dtBeforeMonth);
            DTs.Add(dtCurrentMonth);

            int size = paths.Count;
            for (int i = 0, value = pb_loading.Value + 1; i < size; i++, value++)
            {
                updateProgress(this, value, max);
                int row = 0;
                int col = 0;
                int tStart = 0;
                int tEnd = 0;
                try
                {
                    //1.엑셀열기 및 설정
                    xlapp = new Excel.Application();
                    xlapp.Visible = false;
                    xlapp.UserControl = false;
                    xlapp.DisplayAlerts = false;
                    xlapp.Interactive = true;

                    //읽어올 Excel파일의 경로
                    wb = (Excel.Workbook)xlapp.Workbooks.Open(paths[i]);
                    ws1 = (Excel.Worksheet)wb.Sheets[1];
                    Excel.Range last =
                        ws1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                    row = last.Row;//데이터의 열 길이
                    col = last.Column;//데이터의 행 길이
                    tStart = System.Environment.TickCount;
                    DataTable dt = DTs[i];
                    setDtCols(ref dt, startRowData, col, ws1);//dt에 컬럼 추가
                    if (dt.Columns.Count < 2)
                    {
                        startRowData++;
                        setDtCols(ref dt, startRowData, col, ws1);//dt에 컬럼 추가
                    }
                    setDtRows(dt, row, col, ws1);
                    tEnd = System.Environment.TickCount;
                    Console.WriteLine("setDtRows 수행시간 : " + (tEnd - tStart));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : " + ex.Message);
                }
                finally
                {
                    wb.Close(true);
                    xlapp.Quit();
                    releaseObject(ws1);
                    releaseObject(wb);
                    releaseObject(xlapp);
                    GC.Collect();
                }
            }

            //사업자번호, 단말기ID로 검색
            List<int> closedIdx = new List<int>();
            List<int> newOpenIdx = new List<int>();
            //폐업
            size = dtBeforeMonth.Rows.Count;
            for (int i = 0, value = pb_loading.Value; i < size; i++, value++)
            {
                updateProgress(this, i, max);
                string bn = dtBeforeMonth.Rows[i]["사업자번호"].ToString();
                string slct = string.Format("사업자번호='{0}'", bn);

                DataRow[] stillOpen = dtCurrentMonth.Select(slct);
                if (stillOpen.Length < 1)//폐업 또는 사용종료
                {
                    closedIdx.Add(i);
                    Console.WriteLine((i) + " / " + dtBeforeMonth.Rows[i][1].ToString() + " / " + dtBeforeMonth.Rows[i][2].ToString());
                }
            }
            highlightRow(closedIdx, "yellow", paths[0]);

            //신규
            size = dtCurrentMonth.Rows.Count;
            for (int i = 0, value = pb_loading.Value; i < size; i++, value++)
            {
                updateProgress(this, value, max);

                string bn = dtCurrentMonth.Rows[i]["사업자번호"].ToString();
                string slct = string.Format("사업자번호='{0}'", bn);
                DataRow[] stillOpen = dtBeforeMonth.Select(slct);
                if (stillOpen.Length < 1)//폐업 또는 사용종료
                {
                    newOpenIdx.Add(i);
                    Console.WriteLine((i) + " / " + dtCurrentMonth.Rows[i][1].ToString() + " / " + dtCurrentMonth.Rows[i][2].ToString());
                }
            }

            highlightRow(newOpenIdx, "blue", paths[1]);
            if (pb_loading.Value < max / 2)
            {
                updateProgress(this, max / 2, max);
            }
            if (cb_isMerge.Checked == true)
            {
                mergeDT();
            }
            updateProgress(this, max, max);
            DialogResult rst = MessageBox.Show("완료", "전월/당월 변동사항 확인", MessageBoxButtons.OK);
            if (rst == DialogResult.OK)
            {
                this.Invoke(
                (System.Action)(() => {
                    this.cb_isMerge.Checked = false;
                    updateProgress(this, 0, max);
                }));
            }
        }

        void highlightRow(List<int> workingRow, string color, string path)
        {
            Excel.XlRgbColor background = Excel.XlRgbColor.rgbWhite;
            switch (color.ToLower())
            {
                case "yellow":
                    background = Excel.XlRgbColor.rgbYellow;
                    break;
                case "blue":
                    background = Excel.XlRgbColor.rgbAqua;
                    break;
            }

            //엑셀오픈
            Excel.Application xlapp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws1 = null;

            //1.엑셀열기 및 설정
            xlapp = new Excel.Application();
            xlapp.Visible = false;
            xlapp.UserControl = false;
            xlapp.DisplayAlerts = false;
            xlapp.Interactive = true;
            try
            {
                //읽어올 Excel파일의 경로
                wb = (Excel.Workbook)xlapp.Workbooks.Open(path);
                ws1 = (Excel.Worksheet)wb.Sheets[1];
                Excel.Range last =
                    ws1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                int row = last.Row;//데이터의 열 길이
                int col = last.Column;//데이터의 행 길이
                int size = workingRow.Count;

                for (int i = 0; i < size; i++)
                {
                    Excel.Range selectRow = ws1.get_Range(
                        (Excel.Range)ws1.Cells[startRowData + workingRow[i], 1],
                        (Excel.Range)ws1.Cells[startRowData + workingRow[i], col]);
                    selectRow.Interior.Color = background;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
            finally
            {
                wb.Close(true);
                xlapp.Quit();
                releaseObject(ws1);
                releaseObject(wb);
                releaseObject(xlapp);
                GC.Collect();
            }
        }

        //컬럼값 가져오기
        void setDtCols(ref DataTable dt, int rowCnt, int colLen, Excel.Worksheet ws1)
        {
            //excel에서 header데이터 가져오기
            Excel.Range colunm =
                ws1.get_Range((Excel.Range)ws1.Cells[rowCnt, 1],
                              (Excel.Range)ws1.Cells[rowCnt, colLen]);
            var it = ((IEnumerable)colunm.Value).GetEnumerator();
            while (it.MoveNext())//Excel.Range 범위 순회
            {
                if (it.Current == null)
                {
                    break;
                }
                dt.Columns.Add(it.Current.ToString());
            }
            if (dt.Columns.Count < 2)
            {
                dt.Reset();
            }
        }

        //row값 가져오기
        void setDtRows(DataTable dt, int rowLen, int colLen, Excel.Worksheet ws1)
        {
            Excel.Range excelRow = ws1.get_Range(
                    (Excel.Range)ws1.Cells[startRowData, 1],
                    (Excel.Range)ws1.Cells[startRowData + rowLen + 1, colLen]);

            object[,] obj = (object[,])excelRow.Value;
            int i = 0, j = 0;
            for (i = 1; i <= rowLen; i++)
            {
                List<string> list = new List<string>();
                try
                {
                    for (j = 1; j < colLen + 1; j++)
                    {
                        list.Add(obj[i, j].ToString());
                    }
                    dt.Rows.Add(list.ToArray());
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : " + ex.Message + " / " + list.Count + " , " + i + " , " + j + " , " + string.Join(",", list.ToArray()));
                }
            }

        }

        void mergeDT()
        {
            int len = 2;
            Excel.Application[] xlapp = new Excel.Application[len];
            Excel.Workbook[] wb = new Excel.Workbook[len];
            Excel.Worksheet[] ws = new Excel.Worksheet[len];

            for (int i = 0; i < len; i++)
            {
                //1.엑셀열기 및 설정
                xlapp[i] = new Excel.Application();
                xlapp[i].Visible = false;
                xlapp[i].UserControl = false;
                xlapp[i].DisplayAlerts = false;
                xlapp[i].Interactive = true;
                wb[i] = (Excel.Workbook)xlapp[i].Workbooks.Open(paths[i]);
                ws[i] = (Excel.Worksheet)wb[i].Sheets[1];
            }

            try
            {
                Excel.Range src = ws[1].UsedRange;
                src.Copy(Type.Missing);
                Excel.Range origin =
                    ws[0].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int startCol = DTs[0].Columns.Count + 2;

                Excel.Range dest = ws[0].get_Range(
                    (Excel.Range)(ws[0].Cells[1, startCol]),
                    (Excel.Range)(ws[0].Cells[1 + src.Rows.Count, startCol + src.Columns.Count]));
                dest.Select();
                ws[0].Paste(Type.Missing, Type.Missing);
                wb[0].SaveAs(paths[0], Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, false, false,
                      Excel.XlSaveAsAccessMode.xlExclusive, false, false,
                      null, null, null);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                for (int i = 0; i < len; i++)
                {
                    wb[i].Close(true);
                    xlapp[i].Quit();
                    releaseObject(ws[i]);
                    releaseObject(wb[i]);
                    releaseObject(xlapp[i]);
                }
                GC.Collect();
            }
        }

        void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                Console.WriteLine(e.ToString());
            }
        }

        void updateProgress(Control ctl, object value, object max)
        {
            int iValue = int.Parse(value.ToString());
            int iMax = int.Parse(max.ToString());
            if (ctl.InvokeRequired)
            {
                myDelegate dl = new myDelegate(updateProgress);
                ctl.Invoke(dl, ctl, value, max);
            }
            else
            {
                if (pb_loading.Maximum != iMax)
                {
                    pb_loading.Maximum = iMax;
                }
                if (pb_loading.Value <= iMax)
                {
                    pb_loading.Value = iValue;
                }
            }
        }

    }
}
