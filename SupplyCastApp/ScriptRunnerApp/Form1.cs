using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ScriptRunnerApp
{
    public partial class Form1 : Form
    {
        private Button btnHistory;
        private Button btnForecast;
        private Button btnOrders;
        private Button btnSave;
        private DataGridView dgvOrders;

        public Form1()
        {
            InitializeComponent();

            // Buttons
            btnHistory = new Button { Text = "Generate History", Left = 20, Top = 20, Width = 150 };
            btnForecast = new Button { Text = "Run Forecast", Left = 20, Top = 60, Width = 150 };
            btnOrders = new Button { Text = "Load Orders", Left = 20, Top = 100, Width = 150 };
            btnSave = new Button { Text = "Save Changes", Left = 20, Top = 140, Width = 150 };

            btnHistory.Click += (s, e) => RunScript("python HistoryGenerator.py");
            btnForecast.Click += (s, e) => RunScript("python Forecast.py");
            btnOrders.Click += (s, e) => LoadOrders();
            btnSave.Click += (s, e) => SaveOrders();

            Controls.Add(btnHistory);
            Controls.Add(btnForecast);
            Controls.Add(btnOrders);
            Controls.Add(btnSave);

            // Orders table
            dgvOrders = new DataGridView
            {
                Left = 200,
                Top = 20,
                Width = 600,
                Height = 400,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            Controls.Add(dgvOrders);
        }

        private void RunScript(string command)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", "/c " + command)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    MessageBox.Show(string.IsNullOrEmpty(error) ? output : error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error running script: " + ex.Message);
            }
        }

        private void LoadOrders()
        {
            string filePath = "Orders.xlsx";

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Orders.xlsx not found.");
                return;
            }

            DataTable dt = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // first sheet
                bool firstRow = true;

                foreach (var row in worksheet.RowsUsed())
                {
                    if (firstRow)
                    {
                        foreach (var cell in row.Cells())
                            dt.Columns.Add(cell.Value.ToString());
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add(row.Cells().Select(c => c.Value.ToString()).ToArray());
                    }
                }
            }

            dgvOrders.DataSource = dt;

            // Only allow editing of User Adjustment and Final Order
            foreach (DataGridViewColumn col in dgvOrders.Columns)
            {
                if (col.Name == "User Adjustment" || col.Name == "Final Order")
                {
                    col.ReadOnly = false;
                    col.DefaultCellStyle.BackColor = System.Drawing.Color.LightYellow; // highlight editable
                }
                else
                {
                    col.ReadOnly = true;
                }
            }
        }

        private void SaveOrders()
        {
            string filePath = "Orders.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Orders");

                // Write headers
                for (int i = 0; i < dgvOrders.Columns.Count; i++)
                    worksheet.Cell(1, i + 1).Value = dgvOrders.Columns[i].HeaderText;

                // Write rows
                for (int r = 0; r < dgvOrders.Rows.Count; r++)
                {
                    for (int c = 0; c < dgvOrders.Columns.Count; c++)
                    {
                        object cellValue = dgvOrders.Rows[r].Cells[c].Value;
                        worksheet.Cell(r + 2, c + 1).Value = cellValue != null ? cellValue.ToString() : string.Empty;
                    }
                }

                workbook.SaveAs(filePath);
            }

            MessageBox.Show("Orders.xlsx updated successfully!");
        }

    }
}
