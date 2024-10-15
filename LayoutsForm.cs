using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ProcesadorTxt
{
    public class LayoutsForm : Form
    {
        private DataGridView dataGridView;
        private Button btnCargarArchivo;
        private Button btnExportarExcel;
        private Button btnRegresar;
        private DataTable dataTable;

        public LayoutsForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.dataGridView = new DataGridView();
            this.btnCargarArchivo = new Button();
            this.btnExportarExcel = new Button();
            this.btnRegresar = new Button();
            this.dataTable = new DataTable();

            // Configuración del DataGridView
            this.dataGridView.Location = new System.Drawing.Point(20, 20);
            this.dataGridView.Size = new System.Drawing.Size(750, 400);

            // Configuración del botón de cargar archivo
            this.btnCargarArchivo.Text = "Cargar Archivo";
            this.btnCargarArchivo.Location = new System.Drawing.Point(20, 440);
            this.btnCargarArchivo.Click += new EventHandler(this.BtnCargarArchivo_Click);

            // Configuración del botón de exportar a Excel
            this.btnExportarExcel.Text = "Exportar a Excel";
            this.btnExportarExcel.Location = new System.Drawing.Point(150, 440);
            this.btnExportarExcel.Click += new EventHandler(this.BtnExportarExcel_Click);

            // Configuración del botón de regresar
            this.btnRegresar.Text = "Regresar";
            this.btnRegresar.Location = new System.Drawing.Point(280, 440);
            this.btnRegresar.Click += new EventHandler(this.BtnRegresar_Click);

            // Añadir controles al formulario
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.btnCargarArchivo);
            this.Controls.Add(this.btnExportarExcel);
            this.Controls.Add(this.btnRegresar);

            // Configuración del formulario
            this.Text = "Layouts Processor";
            this.Size = new System.Drawing.Size(800, 600);
        }

        private void BtnCargarArchivo_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog.FileName;
                CargarDatosDesdeExcel(path);
            }
        }

        private void CargarDatosDesdeExcel(string path)
        {
            // Registrar proveedor de codificación para páginas de códigos, necesario para codificaciones como Windows-1252
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            // Usar EPPlus para leer el archivo Excel
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                // var worksheet = package.Workbook.Worksheets[0];  // Asumiendo que la hoja está en la primera posición
                var worksheet = package.Workbook.Worksheets[1];

                // Configurar columnas en el DataTable
                dataTable.Columns.Add("GRUPO", typeof(string));
                dataTable.Columns.Add("GENERICO", typeof(string));
                dataTable.Columns.Add("ESPECIFICADOR", typeof(string));
                dataTable.Columns.Add("DIFERENCIADOR", typeof(string));
                dataTable.Columns.Add("VARIANTE", typeof(string));
                dataTable.Columns.Add("RFC_PROVEEDOR", typeof(string));
                dataTable.Columns.Add("LOTE", typeof(string));
                dataTable.Columns.Add("ESTADO", typeof(int));
                dataTable.Columns.Add("CSUSPENSIVO", typeof(string));
                dataTable.Columns.Add("LINEA", typeof(string));
                dataTable.Columns.Add("LOCALIDAD", typeof(string));
                dataTable.Columns.Add("CANT_INV", typeof(long));
                dataTable.Columns.Add("FECHA_CAD", typeof(DateTime));
                dataTable.Columns.Add("FECHA_FAB", typeof(DateTime));
                dataTable.Columns.Add("FECHA_REC", typeof(DateTime));
                dataTable.Columns.Add("NO_ALTA", typeof(int));
                dataTable.Columns.Add("ESTADO_ANTERIOR", typeof(string));

                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++) // Asumiendo que la primera fila es el encabezado
                {
                    DataRow newRow = dataTable.NewRow();
                    string fuente = worksheet.Cells[row, 10].Text;
                    if (!fuente.ToUpper().Contains("IMSS-BIENESTAR"))
                    {
                        // si la fuente no es IMSS-BIENESTAR, saltar la fila
                        continue;
                    }

                    string claveCNIS = worksheet.Cells[row, 5].Text;

                    // dividir claveCNIS en 4 y meterla en un arreglo
                    string[] claveCNISArray = claveCNIS.Split('.');

                    newRow["GRUPO"] = claveCNISArray[0];
                    newRow["GENERICO"] = claveCNISArray[1];
                    newRow["ESPECIFICADOR"] = claveCNISArray[2];
                    newRow["DIFERENCIADOR"] = claveCNISArray.Length > 3 ? claveCNISArray[3] : "00";
                    newRow["VARIANTE"] = "00";
                    newRow["RFC_PROVEEDOR"] = "XXXX-XXXXXX-XXX";
                    newRow["LOTE"] = worksheet.Cells[row, 8].Text;
                    newRow["ESTADO"] = 1;
                    newRow["CSUSPENSIVO"] = "0";
                    newRow["LINEA"] = "000";
                    newRow["LOCALIDAD"] = "00000000";
                    newRow["CANT_INV"] = long.Parse(worksheet.Cells[row, 11].Text, NumberStyles.AllowThousands);
                    newRow["FECHA_CAD"] = DateTime.TryParse(worksheet.Cells[row, 9].Text, out DateTime fechaCaducidad) ? fechaCaducidad : new DateTime(2025, 12, 31);
                    newRow["FECHA_FAB"] = new DateTime(2024, 1, 1);
                    newRow["FECHA_REC"] = new DateTime(2024, 1, 1);
                    newRow["NO_ALTA"] = 0;
                    newRow["ESTADO_ANTERIOR"] = "0";

                    dataTable.Rows.Add(newRow);
                }

                dataGridView.DataSource = dataTable;
            }
        }

        private void BtnExportarExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ExportarAExcel(saveFileDialog.FileName);
            }
        }

        private void ExportarAExcel(string path)
        {
            // Ordenar los datos por GRUPO, GENERICO, ESPECIFICADOR, DIFERENCIADOR y VARIANTE
            DataView dataView = dataTable.DefaultView;
            dataView.Sort = "GRUPO ASC, GENERICO ASC, ESPECIFICADOR ASC, DIFERENCIADOR ASC, VARIANTE ASC, LOTE ASC";
            DataTable sortedTable = dataView.ToTable();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Layout BCN");

                // Agregar encabezados
                for (int i = 0; i < sortedTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = sortedTable.Columns[i].ColumnName;
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }

                // Rellenar las filas con los datos
                for (int row = 0; row < sortedTable.Rows.Count; row++)
                {
                    for (int col = 0; col < sortedTable.Columns.Count; col++)
                    {
                        var cell = worksheet.Cells[row + 2, col + 1];
                        if (sortedTable.Columns[col].DataType == typeof(DateTime))
                        {
                            // Si es una columna de fecha, usar el formato MM/dd/yyyy HH:mm:ss
                            cell.Style.Numberformat.Format = "MM/dd/yyyy HH:mm:ss";
                            DateTime fecha = Convert.ToDateTime(sortedTable.Rows[row][col]);
                            cell.Value = fecha.ToString("MM/dd/yyyy HH:mm:ss");
                        }
                        else
                        {
                            // Para otros tipos de datos (ej. enteros, decimales)
                            cell.Value = sortedTable.Rows[row][col];
                        }
                    }
                }

                // Autoajustar las columnas
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Guardar el archivo
                FileInfo fi = new FileInfo(path);
                package.SaveAs(fi);

                MessageBox.Show("Archivo Layout generado con éxito con el formato de template.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnRegresar_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 mainForm = new Form1();
            mainForm.Show();
        }
    }
}
