﻿using System;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace Highlighter
{
    delegate void myDelegate(int i, int max);
    public partial class Form1 : Form
    {
        int max = 10000;
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
            if(dialog.ShowDialog() == DialogResult.OK)
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
            if(txt_before.Text.Equals("") || txt_current.Text.Equals("")){
                MessageBox.Show("엑셀파일을 선택 해 주세요.");
                return;
            }
            pb_loading.Maximum = max;
            pb_loading.Value = 0;
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
            //엑셀오픈
            Excel.Application xlapp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws1 = null;

            List<string> paths = new List<string>();
            paths.Add(txt_before.Text.Trim());
            paths.Add(txt_current.Text.Trim());
            DataTable dtBeforeMonth = new DataTable();
            DataTable dtCurrentMonth = new DataTable();
            List<DataTable> DTs = new List<DataTable>();
            DTs.Add(dtBeforeMonth);
            DTs.Add(dtCurrentMonth);
            
            int size = paths.Count;
            for (int i = 0; i < size; i++)
            {
                updateProgress(i, max);
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
                    setDtCols(DTs[i], col, ws1);//dt에 컬럼 추가
                    setDtRows(DTs[i], row, col, ws1);
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
            List<string> closedIdx = new List<string>();
            List<string> newOpenIdx = new List<string>();
            //폐업
            size = dtBeforeMonth.Rows.Count;
            for (int i = 0, value = pb_loading.Value; i < size; i++,value++)
            {
                string bn = dtBeforeMonth.Rows[i]["사업자번호"].ToString();
                string slct = string.Format("사업자번호='{0}'", bn);
                
                DataRow[] stillOpen = dtCurrentMonth.Select(slct);
                if (stillOpen.Length < 1)//폐업 또는 사용종료
                {
                    closedIdx.Add(dtBeforeMonth.Rows[i][0].ToString());
                }
                updateProgress(value, max);
            }
            highlightRow(closedIdx, "yellow", paths[0]);

            //신규
            size = dtCurrentMonth.Rows.Count;
            for (int i = 0, value=pb_loading.Value; i < size; i++, value++)
            {
                string bn = dtCurrentMonth.Rows[i]["사업자번호"].ToString();
                string slct = string.Format("사업자번호='{0}'", bn);
                DataRow[] stillOpen = dtBeforeMonth.Select(slct);
                if (stillOpen.Length < 1)//폐업 또는 사용종료
                {
                    newOpenIdx.Add(dtCurrentMonth.Rows[i][0].ToString());
                }
                updateProgress(value, max);
            }
            highlightRow(newOpenIdx, "blue", paths[1]);
            updateProgress(max, max);
            DialogResult rst = MessageBox.Show("완료","폐업/신규 가맹점 확인", MessageBoxButtons.OK);
            if (rst == DialogResult.OK)
            {
                pb_loading.Value = 0;
            }
        }

        void highlightRow(List<string> workingRow, string color, string path)
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
                    Console.WriteLine("idx : " + workingRow[i]);
                    int rowIdx = int.Parse(workingRow[i]) + 1;
                    Excel.Range selectRow = ws1.get_Range(
                        (Excel.Range)ws1.Cells[rowIdx, 1],
                        (Excel.Range)ws1.Cells[rowIdx, col]);
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
        void setDtCols(DataTable dt, int colLen, Excel.Worksheet ws1)
        {
            //dt 컬럼 추가
            if (dt.Columns.Count < 1)
            {
                //excel에서 header데이터 가져오기
                Excel.Range colunm =
                    ws1.get_Range((Excel.Range)ws1.Cells[1, 1],
                                  (Excel.Range)ws1.Cells[1, colLen]);
                var it = ((IEnumerable)colunm.Value).GetEnumerator();
                while (it.MoveNext())//Excel.Range 범위 순회
                {
                    dt.Columns.Add(it.Current.ToString());
                }
            }
        }

        //row값 가져오기
        void setDtRows(DataTable dt, int rowLen, int colLen, Excel.Worksheet ws1)
        {
            Excel.Range excelRow = ws1.get_Range(
                    (Excel.Range)ws1.Cells[2, 1],
                    (Excel.Range)ws1.Cells[rowLen + 1, colLen]);
            object[,] obj = (object[,])excelRow.Value;
            int i = 0, j = 0;

            for (i = 1; i < rowLen - 1; i++)
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
            }
        }
        
        void updateProgress(int value, int max)
        {
            if(pb_loading.Maximum != max)
            {
                pb_loading.Maximum = max;
            }
            if(pb_loading.Value <= max)
            {
                pb_loading.Value = value;
            }
        }
    }
}
