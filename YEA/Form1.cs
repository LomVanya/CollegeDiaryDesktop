using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data;
using OfficeOpenXml;
using System.IO;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Application = Microsoft.Office.Interop.Excel.Application;
using DocumentFormat.OpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using Table = Microsoft.Office.Interop.Word.Table;
using Row = Microsoft.Office.Interop.Word.Row;
using Cell = Microsoft.Office.Interop.Word.Cell;
using Selection = Microsoft.Office.Interop.Word.Selection;
using Range = Microsoft.Office.Interop.Word.Range;
using PageSetup = Microsoft.Office.Interop.Word.PageSetup;
using System.Diagnostics;

namespace YEA
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }



        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }


        //�������� excel
        private void button1_Click(object sender, EventArgs e)
        {
            
            // �������� ������ ������ ����� Excel.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // ������� ������ Excel � ������� ��������� ����.
                var excelApp = new Application();
                var workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                var worksheet = (Worksheet)workbook.Worksheets[1];

                // �������� �������� �� ��������� �����.
                object[,] values = worksheet.UsedRange.Value;
                int rowCount = values.GetLength(0);
                int columnCount = 4;

                // ��������� DataTable ���������� �� Excel.
                System.Data.DataTable dataTable = new System.Data.DataTable();
                for (int i = 1; i <= columnCount; i++)
                {
                    string columnName = (values[1, i] != null) ? values[1, i].ToString() : "";
                    dataTable.Columns.Add(columnName);
                }
                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        dataRow[j - 1] = values[i, j];
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // ��������� ������� ListView ���������� �� DataTable.
                listView1.Clear();
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    listView1.Columns.Add(dataTable.Columns[i].ColumnName);
                }
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    ListViewItem item = new ListViewItem(dataTable.Rows[i][0].ToString());
                    for (int j = 1; j < dataTable.Columns.Count; j++)
                    {
                        item.SubItems.Add(dataTable.Rows[i][j].ToString());
                    }
                    listView1.Items.Add(item);
                }

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    string columnName = dataTable.Columns[i].ColumnName;
                    int columnWidth = TextRenderer.MeasureText(columnName, listView1.Font).Width;
                    for (int j = 0; j < dataTable.Rows.Count; j++)
                    {
                        string cellValue = dataTable.Rows[j][i].ToString();
                        int cellWidth = TextRenderer.MeasureText(cellValue, listView1.Font).Width;
                        if (cellWidth > columnWidth)
                        {
                            columnWidth = cellWidth;
                        }
                    }
                    listView1.Columns[i].Width = columnWidth;
                }

                for (int i = listView1.Items.Count - 1; i >= 0; i--)
                {
                    ListViewItem item = listView1.Items[i];
                    if (string.IsNullOrEmpty(item.SubItems[1].Text))
                    {
                        listView1.Items.RemoveAt(i);
                    }
                }

                for (int i = listView1.Items.Count - 1; i >= 0; i--)
                {
                    bool isEmpty = true;
                    foreach (ListViewItem.ListViewSubItem subItem in listView1.Items[i].SubItems)
                    {
                        if (!string.IsNullOrEmpty(subItem.Text))
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        listView1.Items.RemoveAt(i);
                    }
                }

                // ������� ���� Excel � ���������� �������.
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                listView1.Columns[0].Width = 150;

                // ������ ������� ����� �������
                int columnIndex = 1;
                int maxValue = 0;

                foreach (ListViewItem item in listView1.Items)
                {
                    if (item.SubItems.Count > columnIndex)
                    {
                        int value;
                        if (int.TryParse(item.SubItems[columnIndex].Text, out value))
                        {
                            if (value > maxValue)
                            {
                                maxValue = value;
                            }
                        }
                    }
                }

                labelVsego.Text = maxValue.ToString();

                int totalQuantity = 0;
                foreach (ListViewItem item in listView1.Items)
                {
                    int quantity;
                    if (int.TryParse(item.SubItems[3].Text, out quantity))
                    {
                        totalQuantity += quantity;
                    }
                }
                label4.Text = totalQuantity.ToString();
            }
        }


        //������� �� ��������� ��� � �����
        private void button2_Click(object sender, EventArgs e)
        {
            List<ListViewItem> currentState = new List<ListViewItem>();

            // �������� ��� �������� �� ListView � ������
            foreach (ListViewItem item in listView1.Items)
            {
                currentState.Add((ListViewItem)item.Clone());
            }

            // ��������� ������� ��������� � ����
            previousStates.Push(currentState);

            DateTime selectedDate = dateTimePicker1.Value;
            int count = (int)numericUpDown1.Value;
            int dateColumnIndex = 0;
            int quantityColumnIndex = 3;
            int filledCellsCount = 0;
            int rowIndex = 0; // �������� � ������� ������

            while (filledCellsCount < count && rowIndex < listView1.Items.Count)
            {
                ListViewItem item = listView1.Items[rowIndex];

                // ���������, ������ �� ������ "�����"
                if (!string.IsNullOrEmpty(item.SubItems[1].Text))
                {
                    // ���� �� ������, �� ���������, ������ �� ������ "����"
                    if (string.IsNullOrEmpty(item.SubItems[dateColumnIndex].Text))
                    {
                        // ���������, �������� �� ��������� ���� ������������
                        if (selectedDate.DayOfWeek == DayOfWeek.Sunday)
                        {
                            MessageBox.Show("�������� ������ ����, ����������� ����������.");
                            return; // ���������� ���������� ������
                        }

                        if (checkBox1.Checked)
                        {
                            item.SubItems[dateColumnIndex].Text = "�";
                        }
                        else
                        {
                            item.SubItems[dateColumnIndex].Text = selectedDate.ToShortDateString();
                        }
                        item.SubItems[quantityColumnIndex].Text = "1";
                        filledCellsCount++;
                    }
                }

                rowIndex++;
            }


            int totalQuantity = 0;
            foreach (ListViewItem item in listView1.Items)
            {
                int quantity;
                if (int.TryParse(item.SubItems[3].Text, out quantity))
                {
                    totalQuantity += quantity;
                }
            }
            label4.Text = totalQuantity.ToString();
        }


        //������� �� ��������� �� listView2 � listView1
        private void button9_Click(object sender, EventArgs e)
        {
            List<ListViewItem> tempList = new List<ListViewItem>();

            int dateColumnIndex = 0;

            for (int i = listView2.Items.Count - 1; i >= 0; i--)
            {
                ListViewItem item = listView2.Items[i];
                if (!string.IsNullOrEmpty(item.SubItems[dateColumnIndex].Text) && item.SubItems[dateColumnIndex].Text.ToLower() != "�")
                {
                    tempList.Add(item);
                    listView2.Items.RemoveAt(i);
                }
            }

            tempList.Sort((x, y) => DateTime.Parse(x.SubItems[dateColumnIndex].Text).CompareTo(DateTime.Parse(y.SubItems[dateColumnIndex].Text)));

            for (int i = tempList.Count - 1; i >= 0; i--)
            {
                ListViewItem item = tempList[i];
                DateTime itemDate = DateTime.Parse(item.SubItems[dateColumnIndex].Text);
                int insertIndex = GetInsertIndex(listView1, itemDate);
                if (insertIndex == -1)
                {
                    listView1.Items.Add(item);
                }
                else
                {
                    listView1.Items.Insert(insertIndex, item);
                }
            }

            int totalQuantity = 0;
            foreach (ListViewItem item in listView1.Items)
            {
                int quantity;
                if (int.TryParse(item.SubItems[3].Text, out quantity))
                {
                    totalQuantity += quantity;
                }
            }
            label4.Text = totalQuantity.ToString();

        }

        // ��������������� ������� ��� ����������� ������� ������� ������ � ��������������� �������
        private int GetInsertIndex(ListView listView, DateTime itemDate)
        {
            int lastIndex = -1;

            for (int i = 0; i < listView.Items.Count; i++)
            {
                DateTime currentDate;
                if (DateTime.TryParse(listView.Items[i].SubItems[0].Text, out currentDate))
                {
                    if (itemDate < currentDate)
                    {
                        return i;
                    }
                    lastIndex = i;
                }
            }

            return lastIndex + 1;
        }




        private void listView1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        Stack<List<ListViewItem>> previousStates = new Stack<List<ListViewItem>>();



        //������ ������
        private void button5_Click(object sender, EventArgs e)
        {
            if (previousStates.Count > 0)
            {
                // ��������� ��������� ��������� �� �����
                List<ListViewItem> previousState = previousStates.Pop();

                // ������� ListView
                listView1.Items.Clear();

                // ��������� �������� �� ����������� ��������� � ListView
                foreach (ListViewItem item in previousState)
                {
                    listView1.Items.Add(item);
                }

                int totalQuantity = 0;
                foreach (ListViewItem item in listView1.Items)
                {
                    int quantity;
                    if (int.TryParse(item.SubItems[3].Text, out quantity))
                    {
                        totalQuantity += quantity;
                    }
                }
                label4.Text = totalQuantity.ToString();

            }
        }

        //�������� ��������������
        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            // �������� ��������� ������� ������
            ListViewItem selectedItem = listView1.SelectedItems[0];

            // ��������� ���������� ���� ��� �������������� ������
            EditItemDialog dialog = new EditItemDialog(selectedItem.SubItems[0].Text, selectedItem.SubItems[1].Text, selectedItem.SubItems[2].Text, selectedItem.SubItems[3].Text);
            DialogResult result = dialog.ShowDialog();

            
                selectedItem.SubItems[0].Text = dialog.Date;
                selectedItem.SubItems[1].Text = dialog.Number;
                selectedItem.SubItems[2].Text = dialog.Name;
                selectedItem.SubItems[3].Text = dialog.Quantity;
           
        }


        //�������� �������������� ��� ���������
        private void listView2_ItemActivate(object sender, EventArgs e)
        {
            // �������� ��������� ������� ������
            ListViewItem selectedItem = listView2.SelectedItems[0];

            // ��������� ���������� ���� ��� �������������� ������
            EditItemDialog dialog = new EditItemDialog(selectedItem.SubItems[0].Text, selectedItem.SubItems[1].Text, selectedItem.SubItems[2].Text, selectedItem.SubItems[3].Text);
            DialogResult result = dialog.ShowDialog();


            selectedItem.SubItems[0].Text = dialog.Date;
            selectedItem.SubItems[1].Text = dialog.Number;
            selectedItem.SubItems[2].Text = dialog.Name;
            selectedItem.SubItems[3].Text = dialog.Quantity;

        }


        //���������� ���
        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count == 0)
            {
                MessageBox.Show("������ ����!");
            }
            else
            {
                // ������� ����� Excel ��������
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filePath = Path.Combine(desktopPath, "ExportedData.xlsx");
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                // ��������� ��������� ��������
                for (int i = 0; i < listView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = listView1.Columns[i].Text;
                }

                // ��������� �������� �� ListView
                for (int i = 0; i < listView1.Items.Count; i++)
                {
                    ListViewItem item = listView1.Items[i];
                    for (int j = 0; j < item.SubItems.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = item.SubItems[j].Text;
                    }
                }
                for (int i = 1; i <= listView1.Columns.Count; i++)
                {
                    Excel.Range column = worksheet.Columns[i];
                    column.ColumnWidth = 20;
                }



                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
                saveFileDialog1.Title = "��������� �������� Excel";
                saveFileDialog1.ShowDialog();

                if (saveFileDialog1.FileName != "")
                {
                    workbook.SaveAs2(saveFileDialog1.FileName);
                }


                workbook.Close();
                excelApp.Quit();

                MessageBox.Show("�������� ��������!");
            }
        }
        //�������
        private void button6_Click(object sender, EventArgs e)
        {
            textBoxName.Clear();
            textBoxGroup.Clear();
            textBoxSpeciality.Clear();
            textBoxCvalification.Clear();
            textBoxSurname.Clear();
            textBoxOtchestvo.Clear();
            textBoxPrepod.Clear();
            textPrepodTwo.Clear();
            textBoxPraktika.Clear();

        }
        //�������� �������
        private void button4_Click(object sender, EventArgs e)
        {

            Oblozhka obl = new Oblozhka();
            obl.Name = textBoxName.Text;
            obl.Surname = textBoxSurname.Text;
            obl.Otchestvo = textBoxOtchestvo.Text;
            obl.Group = textBoxGroup.Text;
            obl.Cvalification = textBoxCvalification.Text;
            obl.Specialization = textBoxSpeciality.Text;
            obl.Prepod = textBoxPrepod.Text;
            obl.PrepodTwo = textPrepodTwo.Text;
            obl.Praktika = textBoxPraktika.Text;

            if (string.IsNullOrEmpty(obl.Name) || string.IsNullOrEmpty(obl.Surname) || string.IsNullOrEmpty(obl.Otchestvo) || string.IsNullOrEmpty(obl.Group) ||
                string.IsNullOrEmpty(obl.Cvalification) || string.IsNullOrEmpty(obl.Specialization) || string.IsNullOrEmpty(obl.Prepod) || string.IsNullOrEmpty(obl.Praktika))
            {
                MessageBox.Show("��������� ��� ����");
            }
            else
            {
                obl.CreateDocument();
                MessageBox.Show("���������!");
            }


        }




        //������ ������
        private void button7_Click(object sender, EventArgs e)
        {
           
            List<ListViewItem> itemsToHighlight = new List<ListViewItem>();

            foreach (ListViewItem item in listView1.Items)
            {
                if (item.SubItems.Count > 2 && item.SubItems[2].Text.Length > 75)
                {
                    itemsToHighlight.Add(item);
                }
            }


            if (listView1.Items.Count == 0)
            {
                MessageBox.Show("������� ��������� �������!");
            }
            else
            {
                label16.Visible = true;
                progressBar1.Visible = true;

                if (itemsToHighlight.Count > 0)
                {
                    listView1.BeginUpdate();

                    foreach (ListViewItem item in itemsToHighlight)
                    {
                        item.BackColor = System.Drawing.Color.Yellow;
                        item.ForeColor = System.Drawing.Color.Black;
                    }

                    listView1.EndUpdate();

                    MessageBox.Show("������! ����� ������ � ������� '�����' ��������� 75 �������.");
                }
                else
                {
                    progressBar1.PerformStep();
                    Word.Application wordApp = new Word.Application();
                    Document wordDoc = wordApp.Documents.Add();


                    int rowsCount = 1;
                    int columnsCount = 9;

                    float tableWidthInCm = 27.74f;
                    float tableWidthInPoints = tableWidthInCm * 28.3465f;

                    Table table = wordDoc.Tables.Add(wordDoc.Paragraphs[1].Range, rowsCount, columnsCount);
                    table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;



                    table.PreferredWidth = tableWidthInPoints;


                    table.Range.Font.Size = 8f;
                    table.Range.Font.Name = "Times New Roman";

                    foreach (Cell cell in table.Range.Cells)
                    {
                        cell.Range.Font.Size = 8f;
                        cell.Range.Font.Name = "Times New Roman";
                        cell.Range.ParagraphFormat.SpaceAfter = 0;

                    }

                    // ������������� ���������� �������� � ������ �������
                    wordDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                    ///////
                    Word.Section section1 = wordDoc.Sections[1];
                    section1.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.6f);
                    section1.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1.0f);
                    //////


                    table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
                    table.PreferredWidth = wordApp.InchesToPoints(8.5f);

                    // ������������� ��������� ��������
                    table.Rows[1].Cells[1].Range.Text = "� �/�";
                    table.Rows[1].Cells[2].Range.Text = "���� ���������� �����";
                    table.Rows[1].Cells[3].Range.Text = "������������ �����";
                    table.Rows[1].Cells[4].Range.Text = "";
                    table.Rows[1].Cells[5].Range.Text = "���������� �����";
                    table.Rows[1].Cells[6].Range.Text = "������� �� ����������� ������";
                    table.Rows[1].Cells[7].Range.Text = "������� ������������ �������� �� ���������� �����������";
                    table.Rows[1].Cells[8].Range.Text = "������� ������������ �������� �� ����������� *";
                    table.Rows[1].Cells[9].Range.Text = "����������";

                    // ������ ����� �������
                    table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

                    // ������ ���� ������ ��� ������� "������"
                    table.Columns[4].Borders[WdBorderType.wdBorderTop].Color = WdColor.wdColorWhite;
                    table.Columns[4].Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorWhite;

                    for (int i = 2; i <= rowsCount; i++)
                    {
                        table.Rows[i].Cells[4].Range.Borders[WdBorderType.wdBorderTop].Color = WdColor.wdColorWhite;
                        table.Rows[i].Cells[4].Range.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorWhite;
                    }

                    // ������ ���� ���������� ������ ��� �������� �������
                    for (int i = 0; i < listView1.Items.Count; i++)
                    {
                        ListViewItem item = listView1.Items[i];

                        // ��������� ����� ������ � �������
                        Row row = table.Rows.Add();

                        // ��������� ������ � ������ �������
                        row.Cells[1].Range.Text = item.SubItems[1].Text; // �����
                        row.Cells[2].Range.Text = item.SubItems[0].Text; // ����
                        row.Cells[3].Range.Text = item.SubItems[2].Text; // ����

                    }

                    progressBar1.PerformStep();

                    float table�ol1 = 0.8f;
                    float col1WidthInPoints = table�ol1 * 28.3465f;

                    float table�ol2 = 1.84f;
                    float col2WidthInPoints = table�ol2 * 28.3465f;

                    float table�ol3 = 10.81f;
                    float col3WidthInPoints = table�ol3 * 28.3465f;

                    float table�ol4 = 1.73f;
                    float col4WidthInPoints = table�ol4 * 28.3465f;

                    float table�ol5 = 1.95f;
                    float col5WidthInPoints = table�ol5 * 28.3465f;

                    float table�ol6 = 2.07f;
                    float col6WidthInPoints = table�ol6 * 28.3465f;

                    float table�ol7 = 3.37f;
                    float col7WidthInPoints = table�ol7 * 28.3465f;

                    float table�ol8 = 3.29f;
                    float col8WidthInPoints = table�ol8 * 28.3465f;

                    float table�ol9 = 1.88f;
                    float col9WidthInPoints = table�ol9 * 28.3465f;

                    table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    table.Columns[1].Width = col1WidthInPoints;
                    table.Columns[2].Width = col2WidthInPoints;
                    table.Columns[3].Width = col3WidthInPoints;
                    table.Columns[4].Width = col4WidthInPoints;
                    table.Columns[5].Width = col5WidthInPoints;
                    table.Columns[6].Width = col6WidthInPoints;
                    table.Columns[7].Width = col7WidthInPoints;
                    table.Columns[8].Width = col8WidthInPoints;
                    table.Columns[9].Width = col9WidthInPoints;


                    int pageCount = wordDoc.ComputeStatistics(Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages, Type.Missing);


                    int PlusStrokes = 45 * pageCount;

                    int StrokesOfNeed = PlusStrokes - listView1.Items.Count;

                    for (int i = 0; i < StrokesOfNeed; i++)
                    {
                        table.Rows.Add();
                    }

                    int startindex = PlusStrokes - 45 + 2;
                    int endindex = table.Rows.Count + 1;
                    int counter = 0;

                    progressBar1.PerformStep();

                    int CountListMinus = listView1.Items.Count - 1;


                    while (startindex > 1)
                    {


                        for (int i = startindex; i < endindex; i++)
                        {

                            if (counter <= CountListMinus)
                            {
                                Row row = table.Rows[i];
                                ListViewItem item = listView1.Items[counter];
                                row.Cells[5].Range.Text = item.SubItems[3].Text;
                                counter++;
                            }
                            else
                            {
                                break;
                            }

                        }


                        startindex -= 45;
                        endindex -= 45;
                    }


                    table.Rows[1].HeadingFormat = -1;
                    progressBar1.PerformStep();


                    float leftIndent = (float)(16 * 28.34646);
                    float fontSize = 8;

                    foreach (Section section in wordDoc.Sections)
                    {
                        section.PageSetup.DifferentFirstPageHeaderFooter = 0;
                        section.PageSetup.FooterDistance = 28.3f;
                    }


                    foreach (Section section in wordDoc.Sections)
                    {
                        Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        Range footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;



                        footerRange.Font.Name = "Times New Roman";
                        footerRange.Font.Size = fontSize;
                        footerRange.ParagraphFormat.LeftIndent = leftIndent;
                        footerRange.Text = "<*> ����������� � ������ ����������� �������� � �����������";


                        //footerRange.InlineShapes.AddHorizontalLineStandard();

                    }

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "�������� Word (*.docx)|*.docx";
                    saveFileDialog1.Title = "��������� �������� Word";
                    saveFileDialog1.ShowDialog();

                    if (saveFileDialog1.FileName != "")
                    {
                        wordDoc.SaveAs2(saveFileDialog1.FileName);
                    }


                    wordDoc.Close();
                    wordApp.Quit();

                    progressBar1.Value = 20;
                    progressBar1.Visible = false;
                    label16.Visible = false;
                }
            }
        }


        //�������� ��������
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                Oblozhka obl = new Oblozhka();
                obl.SignsForming();
                MessageBox.Show("�������� ��������!");
            }
            catch(Exception ex)
            {
                MessageBox.Show("������: " + ex);
            }
        }

        private void listView2_ItemActivate_1(object sender, EventArgs e)
        {

        }

        //���������� ���������

        private void button10_Click(object sender, EventArgs e)
        {

            if (listView2.Items.Count == 0)
            {
                MessageBox.Show("������ ����!");
            }
            else
            {


                // ������� ����� Excel ��������
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filePath = Path.Combine(desktopPath, "ExportedData.xlsx");
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                // ��������� ��������� ��������
                for (int i = 0; i < listView2.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = listView2.Columns[i].Text;
                }

                // ��������� �������� �� ListView
                for (int i = 0; i < listView2.Items.Count; i++)
                {
                    ListViewItem item = listView2.Items[i];
                    for (int j = 0; j < item.SubItems.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = item.SubItems[j].Text;
                    }
                }
                for (int i = 1; i <= listView2.Columns.Count; i++)
                {
                    Excel.Range column = worksheet.Columns[i];
                    column.ColumnWidth = 20;
                }



                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
                saveFileDialog1.Title = "��������� �������� Excel";
                saveFileDialog1.ShowDialog();

                if (saveFileDialog1.FileName != "")
                {
                    workbook.SaveAs2(saveFileDialog1.FileName);
                }


                workbook.Close();
                excelApp.Quit();

                MessageBox.Show("�������� ��������!");
            }

        }

        //�������� excel ���������
        private void button11_Click(object sender, EventArgs e)
        {

            // �������� ������ ������ ����� Excel.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // ������� ������ Excel � ������� ��������� ����.
                var excelApp = new Application();
                var workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                var worksheet = (Worksheet)workbook.Worksheets[1];

                // �������� �������� �� ��������� �����.
                object[,] values = worksheet.UsedRange.Value;
                int rowCount = values.GetLength(0);
                int columnCount = values.GetLength(1);

                // ��������� DataTable ���������� �� Excel.
                System.Data.DataTable dataTable = new System.Data.DataTable();
                for (int i = 1; i <= columnCount; i++)
                {
                    string columnName = (values[1, i] != null) ? values[1, i].ToString() : "";
                    dataTable.Columns.Add(columnName);
                }
                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        dataRow[j - 1] = values[i, j];
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // ��������� ������� ListView ���������� �� DataTable.
                listView2.Clear();
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    listView2.Columns.Add(dataTable.Columns[i].ColumnName);
                }
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    ListViewItem item = new ListViewItem(dataTable.Rows[i][0].ToString());
                    for (int j = 1; j < dataTable.Columns.Count; j++)
                    {
                        item.SubItems.Add(dataTable.Rows[i][j].ToString());
                    }
                    listView2.Items.Add(item);
                }

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    string columnName = dataTable.Columns[i].ColumnName;
                    int columnWidth = TextRenderer.MeasureText(columnName, listView2.Font).Width;
                    for (int j = 0; j < dataTable.Rows.Count; j++)
                    {
                        string cellValue = dataTable.Rows[j][i].ToString();
                        int cellWidth = TextRenderer.MeasureText(cellValue, listView2.Font).Width;
                        if (cellWidth > columnWidth)
                        {
                            columnWidth = cellWidth;
                        }
                    }
                    listView2.Columns[i].Width = columnWidth;
                }

                for (int i = listView2.Items.Count - 1; i >= 0; i--)
                {
                    ListViewItem item = listView2.Items[i];
                    if (string.IsNullOrEmpty(item.SubItems[1].Text))
                    {
                        listView2.Items.RemoveAt(i);
                    }
                }

                for (int i = listView2.Items.Count - 1; i >= 0; i--)
                {
                    bool isEmpty = true;
                    foreach (ListViewItem.ListViewSubItem subItem in listView2.Items[i].SubItems)
                    {
                        if (!string.IsNullOrEmpty(subItem.Text))
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        listView2.Items.RemoveAt(i);
                    }
                }

                // ������� ���� Excel � ���������� �������.
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                listView2.Columns[0].Width = 150;

                
            }
        }


        //������ � �����
        private void button12_Click(object sender, EventArgs e)
        {

            int count = (int)numericUpDown1.Value;
            int dateColumnIndex = 0;
            int filledCellsCount = 0;
            int rowIndex = 0;

            while (filledCellsCount < count && rowIndex < listView1.Items.Count)
            {
                ListViewItem item = listView1.Items[rowIndex];

                // ���������, ����� �� ������ "����" �������� "�"
                if (item.SubItems[dateColumnIndex].Text.ToLower() == "�")
                {
                    // ������� ������� �� listView1 � ��������� ��� � listView2
                    listView1.Items.RemoveAt(rowIndex);
                    listView2.Items.Add((ListViewItem)item.Clone());
                }
                else
                {
                    rowIndex++;
                }
            }

            int totalQuantity = 0;
            foreach (ListViewItem item in listView1.Items)
            {
                int quantity;
                if (int.TryParse(item.SubItems[3].Text, out quantity))
                {
                    totalQuantity += quantity;
                }
            }
            label4.Text = totalQuantity.ToString();
        }

        //����� ���������

        private void button13_Click(object sender, EventArgs e)
        {
            List<ListViewItem> degi = new List<ListViewItem>();

            bool hasError = false;

            foreach (ListViewItem item in listView2.Items)
            {
                if (item.SubItems[0].Text != "�")
                {
                    hasError = true;
                    break;
                }

                degi.Add(item);
            }

            if (listView2.Items.Count == 0 || textBoxSurname.Text == "")
            {
                MessageBox.Show("��������� ��� ��� ������� �� ���������!");
            }
            else if (hasError)
            {
                MessageBox.Show("����������, �������� �������.");
            }
            else
            {
                string surname = textBoxSurname.Text;

                Otrabotka otr = new Otrabotka(degi, surname);
                ReportVew reportView = new ReportVew(otr);

                reportView.Show();

                degi.Clear();
            }
        }



        //�������� �������
        private void SpravkaButton_Click(object sender, EventArgs e)
        {
            string htmlFilePath = @"Spravka\index.htm";

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = htmlFilePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
            
                MessageBox.Show($"������ ��� �������� HTML-�����: {ex.Message}", "������");
            }

        }
    }
}
