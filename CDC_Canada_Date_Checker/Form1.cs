using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CDC_Canada_Date_Checker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            CheckUserAccess();
        }

        private List<string> allowedUsers = new List<string> { "E151859", "E150265", "HAHM" , "E1859", "E0265" };

        private void CheckUserAccess()
        {
            string currentUsername = Environment.UserName.ToUpper(); // Lấy username của Windows

            if (!allowedUsers.Contains(currentUsername)) // Nếu user không nằm trong danh sách
            {
                DisableAllControls(this);
                MessageBox.Show("Access Denied!", "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DisableAllControls(Control parent)
        {
            foreach (Control control in parent.Controls)
            {
                control.Enabled = false; // Vô hiệu hóa tất cả control
                if (control.HasChildren)
                {
                    DisableAllControls(control); // Đệ quy nếu có control con
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != null)
            {
                LoadDataFromExcel(textBox1.Text);

                CheckFile();
                Invoke(new Action(() =>
                {
                    label1.Text = $"Formating table... Please wait";
                    this.Refresh();
                }));
                HighlightErrorCells(dataGridView1);
                
                Invoke(new Action(() =>
                {
                    label1.Text = $"--";
                    this.Refresh();
                }));
                MessageBox.Show("Done!!");
            }
        }

        private void CheckFile()
        {
            string mergedFilePath = Path.GetDirectoryName(textBox1.Text);
            int totalRows = dataGridView1.Rows.Count;
            progressBar1.Maximum = totalRows;
            int processedRows = 0;
            try
            {
                try
                {
                    // Duyệt qua từng dòng trong dataGridView1
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        // Cập nhật tiến trình
                        processedRows++;
                        int progress = processedRows;
                        Invoke(new Action(() =>
                        {
                            progressBar1.Value = progress; // Cập nhật tiến trình trên ProgressBar
                        }));

                        // Lấy giá trị của cột "File_Name" trong dòng hiện tại
                        string fileName = row.Cells["File_Name"].Value.ToString();

                        // Tạo đường dẫn đầy đủ của tệp tin
                        string filePath = Path.Combine(mergedFilePath, fileName);

                        // Kiểm tra xem tệp tin có tồn tại hay không
                        if (File.Exists(filePath))
                        {
                            // Đọc nội dung của tệp tin
                            row.Cells["DateStatus"].Value = "";
                            // Thực hiện xử lý với nội dung của tệp tin

                            ReadDateFromFile(filePath, row);

                        }
                        else
                        {
                            Invoke(new Action(() =>
                            {
                                label1.Text = $"File not found: {filePath}";
                                this.Refresh();
                            }));
                            row.Cells["Status"].Value = $"File not found: {fileName}";
                            Console.WriteLine("Tệp tin {0} không tồn tại.", filePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Lỗi:\r\n{ex.Message}");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Đã xảy ra lỗi: " + e.Message);
            }
        }

        private void HighlightErrorCells(DataGridView dgv)
        {
            dgv.Columns[0].Visible = false;
            dgv.Columns[1].Visible = false;
            dgv.Columns[2].Visible = false;
            dgv.Columns["Status"].Visible = false;

            // Tạm dừng cập nhật giao diện để tránh giật lag
            //dgv.SuspendLayout();
            dgv.BeginInvoke(new Action(() =>
            {
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (cell.Value != null)
                        {
                            string cellText = cell.Value.ToString().ToLower();

                            // Kiểm tra nếu chứa "error" -> đổi màu đỏ, in đậm, căn giữa
                            //if (cellText.Contains("error"))
                            //{
                            //    cell.Style.BackColor = Color.OrangeRed;
                            //    cell.Style.ForeColor = Color.White;
                            //    cell.Style.Font = new Font(dgv.Font, FontStyle.Bold);
                            //    cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            //}
                            //else
                            //{
                            //    // Giữ màu nền mặc định nếu không phải là lỗi
                            //    cell.Style.BackColor = dgv.DefaultCellStyle.BackColor;
                            //    cell.Style.ForeColor = dgv.DefaultCellStyle.ForeColor;
                            //    cell.Style.Font = dgv.DefaultCellStyle.Font;
                            //    cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
                            //}

                            // Nếu ô chứa \r\n hoặc \t -> bật chế độ xuống dòng
                            cell.Style.WrapMode = cellText.Contains("\r\n") || cellText.Contains("\t")
                                ? DataGridViewTriState.True
                                : DataGridViewTriState.False;

                            // Hiển thị tooltip để xem nội dung đầy đủ khi hover chuột
                            cell.ToolTipText = cellText;
                        }
                    }
                }
            }));

            // Tự động điều chỉnh kích thước cột để vừa với màn hình
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // Tự động co giãn hàng để phù hợp với nội dung
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            // Điều chỉnh độ rộng cột Parallel_Citation theo nội dung, các cột khác sẽ giãn đều
            //foreach (DataGridViewColumn col in dataGridView1.Columns)
            //{
            //    if (col.Name == "Parallel_Citation")
            //    {
            //        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; // Fit theo nội dung
            //    }
            //    else
            //    {
            //        col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // Giãn đều các cột còn lại
            //    }
            //}

            // Cập nhật lại hiển thị sau khi thay đổi
            //dgv.ResumeLayout();
        }

        private void LoadDataFromExcel(string filePath)
        {
            DataTable dataTable = new DataTable();

            // Thêm các cột vào DataTable
            //CaseName
            dataTable.Columns.Add("CaseName");
            dataTable.Columns.Add("File_Name");
            dataTable.Columns.Add("Decision_date");
            dataTable.Columns.Add("DateStatus");
            dataTable.Columns.Add("NameStatus");
            dataTable.Columns.Add("Court");
            dataTable.Columns.Add("Topic_Codes");
            dataTable.Columns.Add("Parallel_Citation");
            dataTable.Columns.Add("Special_Instruction");
            dataTable.Columns.Add("Link");
            dataTable.Columns.Add("Status");

            // Đọc dữ liệu từ Excel và thêm vào DataTable
            try
            {
                FileInfo fileInfo = new FileInfo(filePath);

                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy worksheet đầu tiên

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    int decisionDateColumnIndex = -1;
                    int fileNameColumnIndex = -1;
                    int caseNameColumnIndex = -1;
                    int topicCodeColumnIndex = -1;
                    int courtColumnIndex = -1;
                    int parallelCitationColumnIndex = -1;
                    int specialInstructionColumnIndex = -1;

                    // Tìm cột Decision_Date và File_Name
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue = worksheet.Cells[1, col].Value?.ToString();

                        if (cellValue == "Decision_date")
                        {
                            decisionDateColumnIndex = col;
                        }
                        else if (cellValue == "File_Name")
                        {
                            fileNameColumnIndex = col;
                        }
                        else if (cellValue == "CaseName")
                        {
                            caseNameColumnIndex = col;
                        }
                        else if (cellValue == "Topic_Codes")
                        {
                            topicCodeColumnIndex = col;
                        }
                        else if (cellValue == "Court")
                        {
                            courtColumnIndex = col;
                        }
                        else if (cellValue == "Parallel_Citation")
                        {
                            parallelCitationColumnIndex = col;
                        }
                        else if (cellValue == "Special_Instruction")
                        {
                            specialInstructionColumnIndex = col;
                        }

                        // Nếu đã tìm thấy cả hai cột, thoát khỏi vòng lặp
                        if (decisionDateColumnIndex != -1 && fileNameColumnIndex != -1 && caseNameColumnIndex != -1)
                        {
                            break;
                        }
                    }

                    // Kiểm tra xem đã tìm thấy cả hai cột hay chưa
                    if (decisionDateColumnIndex != -1 && fileNameColumnIndex != -1 && caseNameColumnIndex != -1)
                    {
                        // Lấy giá trị từ dòng 2 trở đi của cột Decision_Date và File_Name
                        for (int row = 2; row <= rowCount; row++)
                        {
                            string decisionDate = worksheet.Cells[row, decisionDateColumnIndex].Value?.ToString();
                            string fileName = worksheet.Cells[row, fileNameColumnIndex].Value?.ToString();
                            string caseName = worksheet.Cells[row, caseNameColumnIndex].Value?.ToString();
                            string court = worksheet.Cells[row, courtColumnIndex].Value?.ToString();
                            string topicCode = worksheet.Cells[row, topicCodeColumnIndex].Value?.ToString();
                            string parallel = worksheet.Cells[row, parallelCitationColumnIndex].Value?.ToString();
                            string special = worksheet.Cells[row, specialInstructionColumnIndex].Value?.ToString();

                            // Thêm hàng vào DataTable
                            DataRow dataRow = dataTable.NewRow();
                            dataRow["Decision_Date"] = decisionDate;
                            dataRow["File_Name"] = fileName;
                            dataRow["CaseName"] = caseName;
                            dataRow["Court"] = court;
                            dataRow["Topic_Codes"] = topicCode;
                            dataRow["Parallel_Citation"] = parallel;
                            dataRow["Special_Instruction"] = special;
                            dataTable.Rows.Add(dataRow);
                        }

                        // Hiển thị DataTable trên dataGridView1
                        dataGridView1.DataSource = dataTable;
                    }
                    else
                    {
                        MessageBox.Show("Không tìm thấy cột Decision_Date hoặc File_Name trong tệp Excel.");
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Đã xảy ra lỗi: " + e.Message);
            }

            dataGridView1.DataSource = dataTable;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.Columns["DateStatus"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns["NameStatus"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.Columns["CaseName"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            AutoResizeRowHeight(dataGridView1);
        }

        private void ReadDateFromFile(string filePath, DataGridViewRow row)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);

                int totalLines = lines.Length;
                int processedLines = 0;

                for (int i = 0; i < lines.Length; i++)
                {

                    // Cập nhật tiến độ vào Label1
                    processedLines++;
                    int progress = (int)((double)processedLines / totalLines * 100);
                    Invoke(new Action(() =>
                    {
                        label1.Text = $"Reading {filePath} - Progress: {progress}%";
                        this.Refresh();
                    }));


                    if (lines[i].Contains("<QL:DECISIONDATE/>"))
                    {
                        // Lấy giá trị của dòng kế tiếp
                        string decisionDate = lines[i + 1].Replace(".", "");
                        DateTime sourceDate = ConvertToDate(DecodeHtmlEntities(decisionDate));

                        if (sourceDate == ConvertToDate(row.Cells["Decision_date"].Value.ToString()))
                        {
                            row.Cells["DateStatus"].Value = "OK";
                        }
                        else
                        {
                            row.Cells["DateStatus"].Value = $"{row.Cells["Decision_date"].Value}\r\nXML ->       {DecodeHtmlEntities(decisionDate)}";
                            row.Cells["Link"].Value = filePath;
                        }
                    }
                    if (lines[i].Contains("<QL:SHORTNAME/>"))
                    {
                        // Lấy giá trị của dòng kế tiếp
                        string shortName = lines[i + 1];
                        if (row.Cells["CaseName"].Value.ToString().Replace("\r\n", "") != DecodeHtmlEntities(shortName))
                        {
                            row.Cells["NameStatus"].Value = $"{row.Cells["CaseName"].Value}\r\nXML ->       {DecodeHtmlEntities(shortName)}";
                            row.Cells["Link"].Value = filePath;
                        }
                        else
                        {
                            row.Cells["NameStatus"].Value = "OK";
                        }    
                    }
                    if (lines[i].Contains("<QL:TOPICCODES/>"))
                    {
                        // Lấy giá trị của dòng kế tiếp
                        string topicCode = lines[i + 1];
                        if (row.Cells["Topic_Codes"].Value.ToString().Replace("\r\n", "") != DecodeHtmlEntities(topicCode))
                        {
                            row.Cells["Topic_Codes"].Value = $"{row.Cells["Topic_Codes"].Value}\r\nXML ->        {topicCode}";
                            row.Cells["Link"].Value = filePath;
                        }
                        else
                        {
                            row.Cells["Topic_Codes"].Value = "OK";
                        }
                    }
                    if (lines[i].Contains("<QL:COURTNAME/>"))
                    {
                        // Lấy giá trị của dòng kế tiếp
                        string courtName = lines[i + 1];
                        if (row.Cells["Court"].Value.ToString().Replace("\r\n", "") != DecodeHtmlEntities(courtName))
                        {
                            row.Cells["Court"].Value = $"{row.Cells["Court"].Value}\r\nXML ->     {DecodeHtmlEntities(courtName)}>";
                            row.Cells["Link"].Value = filePath;
                        }
                        else
                        {
                            row.Cells["Court"].Value = "OK";
                        }
                    }
                    if (lines[i].Contains("<QL:PARCITE/>"))
                    {
                        // Lấy giá trị của dòng kế tiếp
                        string topicCode = lines[i + 1];
                        if (row.Cells["Parallel_Citation"].Value.ToString().Replace("\r\n", "") != DecodeHtmlEntities(topicCode))
                        {
                            row.Cells["Parallel_Citation"].Value = $"{row.Cells["Parallel_Citation"].Value}\r\nXML ->         {DecodeHtmlEntities(topicCode)}";
                            row.Cells["Link"].Value = filePath;
                        }
                        else
                        {
                            row.Cells["Parallel_Citation"].Value = "OK";
                        }

                    }
                    if (lines[i].Contains("<QL:CODERCODES/>"))
                    {
                        // Lấy giá trị của dòng kế tiếp
                        string topicCode = lines[i + 1];
                        if (row.Cells["Special_Instruction"].Value.ToString().Replace("\r\n", "") != DecodeHtmlEntities(topicCode))
                        {
                            row.Cells["Special_Instruction"].Value = $"{row.Cells["Special_Instruction"].Value}\r\nXML ->         {DecodeHtmlEntities(topicCode)}";
                            row.Cells["Link"].Value = filePath;
                        }
                        else {
                            row.Cells["Special_Instruction"].Value = "OK";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Đã xảy ra lỗi: " + e.Message);
            }
        }


        static void AutoResizeRowHeight(DataGridView dataGridView)
        {
            // Tự động điều chỉnh chiều cao của mỗi dòng
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.Height = dataGridView.AutoSizeRowsMode == DataGridViewAutoSizeRowsMode.None ? row.Height : 1;
                row.Height = row.GetPreferredHeight(row.Index, DataGridViewAutoSizeRowMode.AllCells, false);
            }
        }

        static DateTime ConvertToDate(string date)
        {
            DateTime dt;
            string[] dateFormats = { "MMMM d, yyyy", "'le' d MMMM yyyy", "MM/dd/yyyy" };

            if (DateTime.TryParseExact(date, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
            {
                return dt;
            }
            else if (DateTime.TryParseExact(date, dateFormats, new CultureInfo("fr-FR"), DateTimeStyles.None, out dt))
            {
                return dt;
            }
            else
            {
                // Trả về DateTime.MinValue nếu không chuyển đổi thành công
                return DateTime.MinValue;
            }
        }

        static string DecodeHtmlEntities(string encodedString)
        {
            return System.Net.WebUtility.HtmlDecode(encodedString);
        }


        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    if (IsExcelFile(file))
                    {
                        e.Effect = DragDropEffects.Copy;
                        return;
                    }
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                textBox1.Text = files[0];
            }
        }

        private bool IsExcelFile(string filePath)
        {
            string extension = Path.GetExtension(filePath);
            return !string.IsNullOrEmpty(extension) && (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase));
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // Kiểm tra xem sự kiện được kích hoạt từ cột link hay không
            if (e.ColumnIndex >= 0 && e.RowIndex >= 0 && sender is DataGridView dataGridView)
            {
                DataGridViewCell clickedCell = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];

                // Kiểm tra xem cell được double click có phải là cột link hay không
                if (clickedCell is DataGridViewLinkCell linkCell)
                {
                    string cellValue = linkCell.Value.ToString();

                    // Kiểm tra xem giá trị của cell có phải là đường dẫn tới file hay không
                    if (File.Exists(cellValue))
                    {
                        // Mở file
                        Process.Start(cellValue);
                    }
                }
            }
        }
    }
}
