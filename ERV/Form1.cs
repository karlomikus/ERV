using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ERV
{
    public partial class mainForm : Form
    {
        private OleDbConnection con;

        StringFormat strFormat; //Used to format the grid rows.
        ArrayList arrColumnLefts = new ArrayList();//Used to save left coordinates of columns
        ArrayList arrColumnWidths = new ArrayList();//Used to save column widths
        int iCellHeight = 0; //Used to get/set the datagridview cell height
        int iTotalWidth = 0; //
        int iRow = 0;//Used as counter
        bool bFirstPage = false; //Used to check whether we are printing first page
        bool bNewPage = false;// Used to check whether we are printing a new page
        int iHeaderHeight = 0; //Used for the header height

        public mainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.btnPrint.Enabled = false;
            this.tbLocation.Text = "att2000.mdb";
            this.con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.tbLocation.Text);
            this.con.Open();

            string userQuery = "SELECT USERID, Name FROM USERINFO";
            OleDbDataAdapter adapter = new OleDbDataAdapter(userQuery, this.con);
            DataSet ds = new DataSet();
            adapter.Fill(ds);
            this.cbUsers.DataSource = ds.Tables[0];
            this.cbUsers.DisplayMember = "Name";
            this.cbUsers.ValueMember = "USERID";

            // Fill months
            this.comboBox2.DataSource = DateTimeFormatInfo.CurrentInfo.MonthNames.Take(12).ToList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = $"SELECT USERINFO.Name, CHECKINOUT.CHECKTIME, CHECKINOUT.CHECKTYPE FROM(CHECKINOUT INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID) WHERE (CHECKINOUT.USERID = {this.cbUsers.SelectedValue})AND (DatePart('m', CHECKTIME) = {(this.comboBox2.SelectedIndex + 1)})AND (DatePart('yyyy', CHECKTIME) = {DateTime.Now.Year.ToString()})ORDER BY CHECKINOUT.CHECKTIME";

            DataTable WorkTimeTable = new DataTable();
            WorkTimeTable.Columns.Add("Ime radnika");
            WorkTimeTable.Columns.Add("Datum");
            WorkTimeTable.Columns.Add("Vrijeme prijave");
            WorkTimeTable.Columns.Add("Vrijeme odjave");
            WorkTimeTable.Columns.Add("Radno vrijeme");

            var r = new List<WorkerWorkTimeEntry>();
            using (OleDbCommand cmd = new OleDbCommand(query, this.con))
            {
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        r.Add(new WorkerWorkTimeEntry { Name = reader.GetString(0), CheckTime = reader.GetDateTime(1), CheckType = reader.GetString(2) });
                    }
                }
            }

            this.btnPrint.Enabled = r.Count() > 0;

            double calcTotalTime = 0;
            var groupedByDate = r.GroupBy(w => w.CheckTime.ToShortDateString()).Select(grp => grp.ToList()).ToList();
            int i = 0;

            foreach (List<WorkerWorkTimeEntry> timeByDate in groupedByDate)
            {
                var wDate = timeByDate.First().CheckTime;
                DataRow dr = WorkTimeTable.NewRow();
                var checkInTime = new DateTime(wDate.Year, wDate.Month, wDate.Day, 8, 0, 0);
                bool hasCheckedIn = false;
                var checkOutTime = new DateTime(wDate.Year, wDate.Month, wDate.Day, 16, 0, 0);
                bool hasCheckedOut = false;

                if (timeByDate.Where(w => w.CheckType == "I").Count() > 0)
                {
                    hasCheckedIn = true;
                    checkInTime = timeByDate.Where(w => w.CheckType == "I").First().CheckTime;
                }

                if (timeByDate.Where(w => w.CheckType == "O").Count() > 0)
                {
                    hasCheckedOut = true;
                    checkOutTime = timeByDate.Where(w => w.CheckType == "O").Last().CheckTime;
                }

                calcTotalTime += checkOutTime.Subtract(checkInTime).TotalHours;

                dr[0] = timeByDate.First().Name;
                dr[1] = wDate.ToShortDateString();
                dr[2] = checkInTime.ToShortTimeString() + (!hasCheckedIn ? " (!)" : "");
                dr[3] = checkOutTime.ToShortTimeString() + (!hasCheckedOut ? " (!)" : "");
                dr[4] = checkOutTime.Subtract(checkInTime).Hours.ToString("0#") + ":" + checkOutTime.Subtract(checkInTime).Minutes.ToString("0#");

                WorkTimeTable.Rows.Add(dr);
                i++;
            }

            this.totalTime.Text = calcTotalTime.ToString("0.##");
            dgResults.DataSource = WorkTimeTable;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = dgPrintDocument;
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                dgPrintDocument.Print();
            }
        }

        private void dgPrintDocument_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidths.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                // Calculating Total Widths
                iTotalWidth = 0;
                foreach (DataGridViewColumn dgvGridCol in dgResults.Columns)
                {
                    iTotalWidth += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgPrintDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {
                //Set the left margin
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;

                //For the first page to print set the cell width and header height
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in dgResults.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width /
                                       (double)iTotalWidth * (double)iTotalWidth *
                                       ((double)e.MarginBounds.Width / (double)iTotalWidth))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText,
                                    GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;

                        // Save width and height of headres
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed
                while (iRow <= dgResults.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = dgResults.Rows[iRow];
                    //Set the cell height
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allo more rows to print
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            string lHeaderText = $"Evidencija radnog vremena ({this.cbUsers.GetItemText(this.cbUsers.SelectedItem)} - {this.comboBox2.GetItemText(this.comboBox2.SelectedItem)})";
                            //Draw Header
                            e.Graphics.DrawString(lHeaderText, new Font(dgResults.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top -
                                    e.Graphics.MeasureString(lHeaderText, new Font(dgResults.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString();
                            //Draw Date
                            e.Graphics.DrawString(strDate, new Font(dgResults.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(dgResults.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString(lHeaderText, new Font(new Font(dgResults.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            //Draw Columns                 
                            iTopMargin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in dgResults.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        //Draw Columns Contents                
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMargin,
                                            (int)arrColumnWidths[iCount], (float)iCellHeight), strFormat);
                            }
                            //Drawing Cells Borders 
                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMargin, (int)arrColumnWidths[iCount], iCellHeight));

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }

                //If more lines exist, print another page.
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
