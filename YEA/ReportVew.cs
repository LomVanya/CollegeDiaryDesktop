using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace YEA
{
    public partial class ReportVew : Form
    {
        private ReportViewer reportViewer;

        private Otrabotka otr;

 

        public ReportVew(Otrabotka otr) 
        {
           
            InitializeComponent();
            this.otr = otr;
           
        }

        private void ReportVew_Load(object sender, EventArgs e)
        {
            DataSet1 dataSet1 = new DataSet1();
           
            reportViewer = new ReportViewer();
            reportViewer.Dock = DockStyle.Fill;
            this.Controls.Add(reportViewer);

            List<string[]> dataList = new List<string[]>();

            foreach (ListViewItem item in otr.DataCollection)
            {
                string[] rowValues = new string[item.SubItems.Count];

                for (int i = 0; i < item.SubItems.Count; i++)
                {
                    rowValues[i] = item.SubItems[i].Text;
                }

                dataList.Add(rowValues);
            }

            foreach (string[] rowArray in dataList)
            {
                DataRow row = dataSet1.DataTable1.NewRow();
                row["Н"] = rowArray[0];
                row["Номер"] = rowArray[1];
                row["Тема"] = rowArray[2];
                row["Часы"] = rowArray[3]; 
                dataSet1.DataTable1.Rows.Add(row);
            }

            reportViewer.LocalReport.DataSources.Clear();
            reportViewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dataSet1.Tables["DataTable1"]));

            reportViewer.LocalReport.ReportPath = "Report1.rdlc";

            ReportParameter datepar = new ReportParameter("Datee", "Дата: " + DateTime.Now.ToString("dd.MM.yyyy"));
            reportViewer.LocalReport.SetParameters(new ReportParameter[] { datepar });

            ReportParameter nameepar = new ReportParameter("Name", "Учащийся: " + otr.Surname);
            reportViewer.LocalReport.SetParameters(new ReportParameter[] { nameepar });

            ReportParameter kolvpar = new ReportParameter("Kol", "Количество пропусков: " + otr.DataCollection.Count.ToString());
            reportViewer.LocalReport.SetParameters(new ReportParameter[] { kolvpar });

            reportViewer.ZoomMode = ZoomMode.PageWidth; 
            reportViewer.RefreshReport();

        }

    }
}
