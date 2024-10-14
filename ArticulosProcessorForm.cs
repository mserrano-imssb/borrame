using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ProcesadorTxt
{
    public class ArticulosProcessorForm : Form
    {
        private DataGridView dataGridView;
        private Button btnCargarArchivo;
        private Button btnExportarExcel;
        private Button btnRegresar;
        private DataTable dataTable;
        private string[] exclude = ["PZA", "ENV", "EQP", "AMP", "CJA", "JGO", "LTA", "BTE", "F.G", "FCO"];
        private string[] headers = ["INSTITUTO MEXICANO DEL SEGURO SOCIAL",
         "NO REFERENCIADO", "BAJA CALIFORNIA NORT", "REPORTE TOTAL DE ARTICULOS", "PAGINA",
          "ARTICULO     ", "ARTICULO PRESENTACION", "P.U.U."];


        public ArticulosProcessorForm()
        {
            InitializeComponent();
            ConfigurarDataGridView();
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
            this.Text = "Articulos Processor";
            this.Size = new System.Drawing.Size(800, 600);
        }

        private void BtnCargarArchivo_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt",
                Title = "Seleccionar archivo TXT"
            };

            var dialogResult = openFileDialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                ProcesarArchivo(openFileDialog.FileName);
            }
            else
            {
                var result = dialogResult == DialogResult.Abort ? "Abort" : "Cancel";
                MessageBox.Show("Archivo no seleccionado", "Error " + result, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConfigurarDataGridView()
        {
            // Configuración de las columnas
            dataTable.Columns.Add("NUM", typeof(string));
            dataTable.Columns.Add("GPO", typeof(string));
            dataTable.Columns.Add("GEN", typeof(string));
            dataTable.Columns.Add("ESP", typeof(string));
            dataTable.Columns.Add("DIF", typeof(string));
            dataTable.Columns.Add("VAR", typeof(string));
            dataTable.Columns.Add("P.P", typeof(string));
            dataTable.Columns.Add("CBI", typeof(string));
            dataTable.Columns.Add("CLAVE", typeof(string));
            dataTable.Columns.Add("DESCRIPCION", typeof(string));
            dataTable.Columns.Add("U/M", typeof(string));
            dataTable.Columns.Add("CANTIDAD", typeof(decimal));
            dataTable.Columns.Add("TIPO", typeof(string));
            dataTable.Columns.Add("T.A", typeof(string));
            dataTable.Columns.Add("P.U.U", typeof(decimal));
        }

        private void ProcesarArchivo(string filePath)
        {
            string line = "";
            string descripcionAcumulada = "";
            try
            {
                // Registrar proveedor de codificación para páginas de códigos, necesario para codificaciones como Windows-1252
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                // Leer el archivo con la codificación ANSI (Windows-1252)
                string[] lines = File.ReadAllLines(filePath, Encoding.GetEncoding("Windows-1252"));

                foreach (string linea in lines)
                {
                    line = linea.Replace("_", "").Replace("=", "").Trim();

                    /* Si line empieza con una letra, remover exceso de espacios en blanco entre estos, por ejemplo, quitar
                     la laguna de espacio en blanco en cadenas como: "DE CADENA LARGA AL 20 %                               OL" */
                    if (!string.IsNullOrEmpty(descripcionAcumulada) &&
                        !string.IsNullOrWhiteSpace(line) &&
                         char.IsLetter(line[0]))
                    {
                        // Need to trim inbetween spaces by splitting the string with space and then joining it again
                        line = string.Join(" ", line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries));
                    }

                    if (EsLineaValida(line))
                    {
                        if (line.Length > 50)
                        {
                            string[] preColumns = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            List<string> columnsList = new List<string>();
                            var descripcion = "";
                            columnsList.Add(preColumns[0]); // NUM
                            columnsList.Add(preColumns[1]); // GPO
                            columnsList.Add(preColumns[2]); // GEN
                            columnsList.Add(preColumns[3]); // ESP
                            columnsList.Add(preColumns[4]); // DIF
                            columnsList.Add(preColumns[5]); // VAR
                            columnsList.Add(preColumns[6]); // P.P
                            columnsList.Add(preColumns[7]); // CBI
                            string clave = preColumns[1] + "." + preColumns[2] + "." + preColumns[3] + "." + preColumns[4];
                            columnsList.Add(clave);

                            if (preColumns[0] == "25")
                            {
                                var popo = 0;
                            }

                            int i = 8;
                            for (i = 8; i < preColumns.Length &&
                                            descripcion.Length + preColumns[i].Length <= 50 &&
                                            (
                                                 !exclude.Contains(preColumns[i])
                                                 ||
                                                 (
                                                     descripcion.Length + preColumns[i].Length == 50
                                                     &&
                                                     exclude.Contains(preColumns[i])
                                                 )
                                            );
                                 i++)
                            {
                                descripcion += preColumns[i] + " ";
                                if ((line.Contains("DE   CLORURO   DE") ||
                                      line.Contains("PARA  CATETERISMO  VENOSO  CENTRAL  RAD") ||
                                      line.Contains("CONSTAN    DE    CIER") ||
                                      line.Contains("DISPOSITIVO  INTRAUTERINO  T  DE  CO") ||
                                      line.Contains("JERINGA  DESECHABLE  PARA  APLICAR  ") ||
                                      line.Contains("ENDOBRONQUIAL  PARA  INTUBACIÓN  DE  "))
                                && i <= 11)//chicanada
                                {
                                    descripcion += " "; // agrega un espacio al final para hacer bola y evitar que se coma el campo de u/m
                                }
                            }
                            columnsList.Add(descripcion);
                            // agregar el resto de columnas
                            for (int j = i; j < preColumns.Length; j++)
                            {
                                columnsList.Add(preColumns[j]);
                            }

                            var columns = columnsList.ToArray(); // siempre son 14 columnas


                            DataRow row = dataTable.NewRow();
                            row["NUM"] = columns[0].Trim();
                            row["GPO"] = columns[1].Trim();
                            row["GEN"] = columns[2].Trim();
                            row["ESP"] = columns[3].Trim();
                            row["DIF"] = columns[4].Trim();
                            row["VAR"] = columns[5].Trim();
                            row["P.P"] = columns[6].Trim();
                            row["CBI"] = columns[7].Trim();
                            row["CLAVE"] = columns[8].Trim();
                            row["DESCRIPCION"] = columns[9].Trim();
                            row["U/M"] = columns[10].Trim();
                            row["CANTIDAD"] = decimal.Parse(columns[11].Trim());
                            row["TIPO"] = columns[12].Trim();
                            row["T.A"] = columns[13].Trim();
                            row["P.U.U"] = decimal.Parse(columns[14].Trim());

                            // Acumular descripción en caso de múltiples líneas
                            descripcionAcumulada = row["DESCRIPCION"].ToString();
                            dataTable.Rows.Add(row);
                        }
                        else
                        {
                            // Si es la segunda línea de descripción, concatenar con la descripción anterior
                            descripcionAcumulada += line.Trim();
                            dataTable.Rows[dataTable.Rows.Count - 1]["DESCRIPCION"] = descripcionAcumulada;

                            // Limpiar descripción acumulada
                            descripcionAcumulada = "";
                        }
                    }
                }

                dataGridView.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// Metodo para validar si el renglon corresponde a dato de articulo
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private bool EsLineaValida(string line)
        {
            // Validar si al menos uno de los elementos en this.headers esta presente en la línea
            foreach (string header in this.headers)
            {
                if (line.Contains(header))
                {
                    return false;
                }
            }

            // Validar si la línea es una línea válida que contiene datos de artículos
            return !string.IsNullOrWhiteSpace(line) && line.Length > 5 &&
                    (char.IsDigit(line[0]) || char.IsLetter(line[0]));
        }

        private void BtnExportarExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Guardar como Excel"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                ExportarAExcel(saveFileDialog.FileName);
            }
        }

        private void ExportarAExcel(string filePath)
        {
            // necesito ordenar tabla por las primeras 4 columnas
            dataTable.DefaultView.Sort = "GPO, GEN, ESP, DIF, VAR";
            dataTable = dataTable.DefaultView.ToTable();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Datos");

                // Exportar encabezados
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                // Exportar filas
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                ExportarConTemplateAExcel(filePath);
            }
        }

        private void ExportarConTemplateAExcel(string filePath)
        {
            // Crear un nuevo archivo Excel
            using (var package = new ExcelPackage())
            {
                // Añadir una hoja llamada "Articulos"
                var worksheet = package.Workbook.Worksheets.Add("Articulos");

                // Crear el encabezado con el estilo deseado
                string[] header = { "NUM", "GPO", "GEN", "ESP", "DIF", "VAR", "P.P", "CBI", "CLAVE", "DESCRIPCION", "U/M", "CANTIDAD", "TIPO", "T.A", "P.U.U" };

                for (int i = 0; i < header.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = header[i];
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                    worksheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                // Poblar los datos en las filas debajo del encabezado
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                    }
                }

                // Autoajustar las columnas
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Guardar el archivo Excel en la ruta proporcionada
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);

                MessageBox.Show("Catálogo de artículos exportado con éxito a excel.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
