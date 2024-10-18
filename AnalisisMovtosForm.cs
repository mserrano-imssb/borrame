using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ProcesadorTxt
{
    public partial class AnalisisMovtosForm : Form, IFormWithLoadedData
    {
        private DataGridView dataGridView;
        private Button btnCargarArchivo;
        private Button btnExportarExcel;
        private List<MovimientoArticulo> movimientos; // Lista para almacenar todos los registros de movimientos
        private TableLayoutPanel mainLayoutPanel;
        private FlowLayoutPanel buttonPanel;

        public AnalisisMovtosForm()
        {
            InitializeComponent();
            MostrarDatosEnGrid();
        }

        public bool HasDataLoaded()
        {
            // Verificar si el DataTable tiene filas
            return movimientos != null && movimientos.Count > 0;
        }

        private void InitializeComponent()
        {
            this.dataGridView = new DataGridView();
            this.btnCargarArchivo = new Button();
            this.btnExportarExcel = new Button();
            this.movimientos = new List<MovimientoArticulo>();

            // Inicialización del TableLayoutPanel principal
            this.mainLayoutPanel = new TableLayoutPanel();
            this.mainLayoutPanel.ColumnCount = 1;
            this.mainLayoutPanel.RowCount = 2;
            this.mainLayoutPanel.Dock = DockStyle.Fill;
            this.mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 155F)); // 155% para el DataGridView
            this.mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // Fila para botones

            // Configuración del DataGridView
            this.dataGridView.Dock = DockStyle.Fill;
            this.dataGridView.Size = new System.Drawing.Size(750, 400);

            // Configuración del botón de cargar archivo
            this.btnCargarArchivo.Text = "Cargar Archivo";
            this.btnCargarArchivo.Size = new System.Drawing.Size(120, 30);
            this.btnCargarArchivo.Click += new EventHandler(this.BtnCargarArchivo_Click);

            // Configuración del botón de exportar a excel
            this.btnExportarExcel.Text = "Exportar a Excel";
            this.btnExportarExcel.Size = new System.Drawing.Size(120, 30);
            this.btnExportarExcel.Click += new EventHandler(this.BtnExportarExcel_Click);

            // Panel para los botones
            buttonPanel = new FlowLayoutPanel();
            buttonPanel.Dock = DockStyle.Fill;
            buttonPanel.FlowDirection = FlowDirection.LeftToRight; // Botones de izquierda a derecha
            buttonPanel.WrapContents = false;  // Evitar que los botones se apilen en múltiples líneas
            buttonPanel.AutoSize = true; // Permitir que el panel se ajuste al contenido

            // Añadir los botones al buttonPanel
            buttonPanel.Controls.Add(this.btnCargarArchivo);
            buttonPanel.Controls.Add(this.btnExportarExcel);

            // Agregar controles al TableLayoutPanel
            this.mainLayoutPanel.Controls.Add(this.dataGridView, 0, 0);  // Primera fila (DataGridView)
            this.mainLayoutPanel.Controls.Add(buttonPanel, 0, 1);

            this.Controls.Add(this.mainLayoutPanel);

            // Configuración del formulario
            this.Text = "Análisis de Movimientos de Artículos";
            this.Size = new System.Drawing.Size(850, 600);
        }

        private void BtnCargarArchivo_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Text Files|*.txt",
                Title = "Seleccionar archivo de análisis de movimientos con lotes"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ProcesarArchivo(openFileDialog.FileName);
            }
        }

        private void ProcesarArchivo(string filePath)
        {
            // Registrar proveedor de codificación para páginas de códigos, necesario para codificaciones como Windows-1252
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Leer el archivo línea por línea y procesarlo
            var lines = File.ReadAllLines(filePath, Encoding.GetEncoding("Windows-1252"));
            MovimientoArticulo movimientoPivote = null;
            string currentGrupo = null;
            this.movimientos = new List<MovimientoArticulo>();
            btnExportarExcel.Enabled = false;
            var subBucleActivado = false;

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];

                if (EsLineaIgnorada(line))
                    continue; // Ignorar líneas según las pautas

                if (line.Contains("GRUPO:"))
                {
                    if (!subBucleActivado)
                        currentGrupo = ObtenerGrupo(line);
                    continue;
                }

                if (line.Contains("ARTICULO:"))
                {
                    if (!subBucleActivado)
                    {
                        // Proceso de obtención de [ARTICULO], [DESCRIPCION], etc.
                        movimientoPivote = new MovimientoArticulo();
                        movimientoPivote.Grupo = currentGrupo;
                        movimientoPivote.Articulo = ObtenerArticulo(line);
                        movimientoPivote.Descripcion = ObtenerDescripcion(line);
                        movimientoPivote.Uni = ObtenerUni(line);
                        movimientoPivote.CantidadPresentacion = ObtenerCantidadPresentacion(line);
                        movimientoPivote.Tipo = ObtenerTipo(line);
                        movimientoPivote.Puu = ObtenerPuu(line);
                    }
                    continue;
                }

                if (line.Contains("EXISTENCIA INICIAL"))
                {
                    movimientoPivote.ExistenciaInicial = ObtenerExistenciaInicial(line);
                    subBucleActivado = true;
                    continue;
                }

                if (line.Contains("EXISTENCIA FINAL"))
                {
                    // Proceso para existencias finales
                    var existenciaFinal = ObtenerExistenciasFinales(line);
                    ActualizarExistenciasFinales(movimientoPivote.Articulo, existenciaFinal);
                    subBucleActivado = false;
                    continue;
                }

                // Proceso para los movimientos
                if (subBucleActivado) //EsLineaMovimiento(line))
                {
                    // Obtener los movimientos y los datos asociados
                    var movimiento = ProcesarMovimiento(line, lines, movimientoPivote, ref i);
                    movimientos.Add(movimiento);
                }
            }

            // Mostrar los datos en el DataGridView
            MostrarDatosEnGrid();
            // habilitar boton de excel si hay datos
            if (movimientos.Count > 0)
            {
                btnExportarExcel.Enabled = true;
            }
            else
            {
                MessageBox.Show("No se encontraron datos. Verifique que tenga formato correcto.", "Datos no encontrados", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                btnExportarExcel.Enabled = false;
            }
        }

        private void MostrarDatosEnGrid()
        {
            // Vincular la lista de movimientos al DataGridView
            dataGridView.DataSource = movimientos;
        }

        private bool EsLineaIgnorada(string line)
        {
            return string.IsNullOrWhiteSpace(line) ||
                   line.Contains("CLASF. PRESP.") ||
                   line.Contains("IMSS-SAI MODULO DE ALMACENES") ||
                   line.Contains("ALMACÉN ZONA") ||
                   line.Contains("ANALISIS DE MOVIMIENTOS DE ARTICULO") ||
                   line.Contains("DEL ARTICULO") ||
                   line.Contains("TIPO DE MOVIMIENTO") ||
                   line.Contains("PROVEEDOR") ||
                   line.Contains("                                                                                          CANTIDAD") ||
                   line.Contains("--------------------------------------------------------------------------------------------------------------") ||
                   line.Length > 110;
        }

        private bool EsLineaMovimiento(string line)
        {
            return !string.IsNullOrWhiteSpace(line) && line.Length <= 110 && line.Contains("MOVIMIENTO");
        }

        private string ObtenerGrupo(string line)
        {
            // Obtener el grupo a partir de IndexOf("GRUPO:") + 6 para omitir la palabra "GRUPO:"
            return line.Substring(line.IndexOf("GRUPO:") + 6).Trim();
        }

        private string ObtenerArticulo(string line)
        {
            return line.Substring(line.IndexOf("ARTICULO:") + 9, 19).Trim();
        }

        private string ObtenerDescripcion(string line)
        {
            return line.Substring(line.IndexOf("ARTICULO:") + 28, 36).Trim();
        }

        private string ObtenerUni(string line)
        {
            return line.Substring(line.IndexOf("ARTICULO:") + 64, 3).Trim();
        }

        private string ObtenerCantidadPresentacion(string line)
        {
            var parts = line.Substring(67).Trim().Split(' ');
            return parts.Length > 0 ? parts[0] : string.Empty;
        }

        private string ObtenerTipo(string line)
        {
            var parts = line.Substring(67).Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 1 ? parts[1] : string.Empty;
        }

        private string ObtenerPuu(string line)
        {
            var parts = line.Substring(67).Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return parts.Length > 2 ? parts[2] : string.Empty;
        }

        private string ObtenerExistenciaInicial(string line)
        {
            return line.Substring(line.IndexOf("EXISTENCIA INICIAL") + 19).Trim();
        }

        private ExistenciasFinales ObtenerExistenciasFinales(string line)
        {
            // Obtener las existencias finales
            var partes = line.Substring(77).Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
            return new ExistenciasFinales
            {
                Entradas = partes.Length > 0 ? partes[0] : "0",
                Salidas = partes.Length > 1 ? partes[1] : "0",
                Saldos = partes.Length > 2 ? partes[2] : "0"
            };
        }

        private void ActualizarExistenciasFinales(string articulo, ExistenciasFinales existenciaFinal)
        {
            // Buscar todos los registros para el artículo actual y actualizar sus existencias finales
            foreach (var mov in movimientos.Where(m => m.Articulo == articulo && m.ExistenciaFinalEntradas == "0"))
            {
                mov.ExistenciaFinalEntradas = existenciaFinal.Entradas;
                mov.ExistenciaFinalSalidas = existenciaFinal.Salidas;
                mov.ExistenciaFinalSaldos = existenciaFinal.Saldos;
            }
        }

        /// <summary>
        /// Procesa movimientos mediante subbucle
        /// </summary>
        /// <param name="line"></param>
        /// <param name="lines"></param>
        /// <param name="i"></param>
        /// <returns></returns>
        private MovimientoArticulo ProcesarMovimiento(string line, string[] lines, MovimientoArticulo movimientoPivote, ref int i)
        {
            var movimiento = new MovimientoArticulo();

            // Procesar los campos del movimiento
            movimiento.TipoMovimiento = line.Substring(0, 35).Trim();
            movimiento.Documento = line.Substring(35, 10).Trim();
            movimiento.UnidadProveedor = line.Substring(46, 18).Trim();
            movimiento.Fecha = line.Substring(64, 11).Trim();

            var ess = line.Substring(76, 34).Trim();
            var essList = ess.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            movimiento.Entradas = essList.Length > 0 ? essList[0] : "0";
            movimiento.Salidas = essList.Length > 1 ? essList[1] : "0";
            movimiento.Saldos = essList.Length > 2 ? essList[2] : "0";

            // Verificar si la línea siguiente contiene información de proveedor/lote
            if (lines.Length > i + 1 && lines[i + 1].Contains("PROVEEDOR"))
            {
                i += 2; // Saltar a la línea con los datos de proveedor
                var proveedorLine = lines[i];
                movimiento.NombreProveedor = proveedorLine.Substring(0, 62).Trim();
                movimiento.Lote = proveedorLine.Substring(62, 21).Trim();
                movimiento.Caducidad = proveedorLine.Substring(84, 11).Trim();
                movimiento.Cantidad = proveedorLine.Substring(95, 15).Trim();
            }

            // agregar el resto de campos en movimiento desde movimientoPivote
            if (movimientoPivote != null)
            {
                movimiento.ExistenciaInicial = movimientoPivote.ExistenciaInicial;
                movimiento.Grupo = movimientoPivote.Grupo;
                movimiento.Articulo = movimientoPivote.Articulo;
                movimiento.Descripcion = movimientoPivote.Descripcion;
                movimiento.Uni = movimientoPivote.Uni;
                movimiento.CantidadPresentacion = movimientoPivote.CantidadPresentacion;
                movimiento.Tipo = movimientoPivote.Tipo;
                movimiento.Puu = movimientoPivote.Puu;
            }

            return movimiento;
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
            // Registrar proveedor de codificación para páginas de códigos, necesario para codificaciones como Windows-1252
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Movimientos Artículos");

                // Agregar las cabeceras al Excel
                worksheet.Cells[1, 1].Value = "GRUPO";
                worksheet.Cells[1, 2].Value = "ARTICULO";
                worksheet.Cells[1, 3].Value = "DESCRIPCION";
                worksheet.Cells[1, 4].Value = "UNI";
                worksheet.Cells[1, 5].Value = "CANTIDAD PRESENTACION";
                worksheet.Cells[1, 6].Value = "TIPO";
                worksheet.Cells[1, 7].Value = "PUU";
                worksheet.Cells[1, 8].Value = "EXISTENCIA INICIAL";
                worksheet.Cells[1, 9].Value = "TIPO MOVIMIENTO";
                worksheet.Cells[1, 10].Value = "DOCUMENTO";
                worksheet.Cells[1, 11].Value = "UNIDAD/PROVEEDOR";
                worksheet.Cells[1, 12].Value = "FECHA";
                worksheet.Cells[1, 13].Value = "ENTRADAS";
                worksheet.Cells[1, 14].Value = "SALIDAS";
                worksheet.Cells[1, 15].Value = "SALDOS";
                worksheet.Cells[1, 16].Value = "NOMBRE PROVEEDOR";
                worksheet.Cells[1, 17].Value = "LOTE";
                worksheet.Cells[1, 18].Value = "CADUCIDAD";
                worksheet.Cells[1, 19].Value = "CANTIDAD";
                worksheet.Cells[1, 20].Value = "EXISTENCIA FINAL ENTRADAS";
                worksheet.Cells[1, 21].Value = "EXISTENCIA FINAL SALIDAS";
                worksheet.Cells[1, 22].Value = "EXISTENCIA FINAL SALDOS";

                // Rellenar el Excel con los datos de la lista de movimientos
                int row = 2;
                foreach (var mov in movimientos)
                {
                    worksheet.Cells[row, 1].Value = mov.Grupo;
                    worksheet.Cells[row, 2].Value = mov.Articulo;
                    worksheet.Cells[row, 3].Value = mov.Descripcion;
                    worksheet.Cells[row, 4].Value = mov.Uni;
                    // worksheet.Cells[row, 5].Value = mov.CantidadPresentacion;
                    if (decimal.TryParse(mov.CantidadPresentacion, NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, CultureInfo.CurrentCulture, out decimal cantidadPresentacion))
                    {
                        worksheet.Cells[row, 5].Value = cantidadPresentacion; // Convertido a decimal
                        worksheet.Cells[row, 5].Style.Numberformat.Format = "0.000"; // Formato con 3 decimales, sin separador de miles
                    }
                    else
                    {
                        worksheet.Cells[row, 5].Value = mov.Cantidad; // Si no se puede convertir, dejar el valor original
                    }

                    worksheet.Cells[row, 6].Value = mov.Tipo;

                    // Convertir el campo PUU de string a decimal (campo monetario)
                    if (decimal.TryParse(mov.Puu, NumberStyles.Currency | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out decimal puu))
                    {
                        worksheet.Cells[row, 7].Value = puu; // Convertido a decimal
                        worksheet.Cells[row, 7].Style.Numberformat.Format = "$#,##0.00"; // Formato de moneda en Excel
                    }
                    else
                    {
                        worksheet.Cells[row, 7].Value = mov.Puu; // Si no se puede convertir, dejar el valor original
                    }

                    worksheet.Cells[row, 8].Value = long.Parse(mov.ExistenciaInicial, NumberStyles.AllowThousands); // Convertir a long
                    worksheet.Cells[row, 9].Value = mov.TipoMovimiento;
                    worksheet.Cells[row, 10].Value = mov.Documento;
                    worksheet.Cells[row, 11].Value = mov.UnidadProveedor;

                    // worksheet.Cells[row, 12].Value = mov.Fecha;
                    // Convertir fecha si es válida
                    if (TryConvertirFecha(mov.Fecha, out DateTime fecha))
                    {
                        worksheet.Cells[row, 12].Value = fecha;
                        worksheet.Cells[row, 12].Style.Numberformat.Format = "dd/MM/yyyy";
                    }
                    else
                    {
                        worksheet.Cells[row, 12].Value = mov.Fecha; // Dejar el valor original si no se puede convertir
                    }

                    worksheet.Cells[row, 13].Value = long.Parse(mov.Entradas, NumberStyles.AllowThousands); // Convertir a long
                    worksheet.Cells[row, 14].Value = long.Parse(mov.Salidas, NumberStyles.AllowThousands); // Convertir a long
                    worksheet.Cells[row, 15].Value = long.Parse(mov.Saldos, NumberStyles.AllowThousands); // Convertir a long
                    worksheet.Cells[row, 16].Value = mov.NombreProveedor;
                    worksheet.Cells[row, 17].Value = mov.Lote;

                    //worksheet.Cells[row, 18].Value = mov.Caducidad;
                    // Convertir caducidad si es válida
                    if (TryConvertirFecha(mov.Caducidad, out DateTime caducidad))
                    {
                        worksheet.Cells[row, 18].Value = caducidad;
                        worksheet.Cells[row, 18].Style.Numberformat.Format = "dd/MM/yyyy";
                    }
                    else
                    {
                        worksheet.Cells[row, 18].Value = mov.Caducidad; // Dejar el valor original si no se puede convertir
                    }

                    worksheet.Cells[row, 19].Value = long.Parse(mov.Cantidad, NumberStyles.AllowThousands);
                    worksheet.Cells[row, 20].Value = long.Parse(mov.ExistenciaFinalEntradas, NumberStyles.AllowThousands);
                    worksheet.Cells[row, 21].Value = long.Parse(mov.ExistenciaFinalSalidas, NumberStyles.AllowThousands);
                    worksheet.Cells[row, 22].Value = long.Parse(mov.ExistenciaFinalSaldos, NumberStyles.AllowThousands);

                    row++;
                }

                // Guardar el archivo Excel
                package.SaveAs(new FileInfo(filePath));

                MessageBox.Show("Archivo Excel exportado correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Intenta convertir fecha dado que el TXT de movimientos de articulos
        /// está estrictamente con fechas con nombres de mes en español
        /// </summary>
        /// <param name="fechaTexto"></param>
        /// <param name="fecha"></param>
        /// <returns></returns>
        private bool TryConvertirFecha(string fechaTexto, out DateTime fecha)
        {
            fecha = DateTime.MinValue;

            // Diccionario para los meses en español
            var mesesEspanol = new Dictionary<string, int>
                            {
                                { "ENE", 1 }, { "FEB", 2 }, { "MAR", 3 }, { "ABR", 4 },
                                { "MAY", 5 }, { "JUN", 6 }, { "JUL", 7 }, { "AGO", 8 },
                                { "SEP", 9 }, { "OCT", 10 }, { "NOV", 11 }, { "DIC", 12 }
                            };

            // Regex para identificar las fechas en el formato "dd/MMM/yyyy" (ej: 21/ENE/2026)
            var regex = new Regex(@"(\d{2})/([A-Z]{3})/(\d{4})");
            var match = regex.Match(fechaTexto);

            if (match.Success)
            {
                int dia = int.Parse(match.Groups[1].Value);
                string mesTexto = match.Groups[2].Value.ToUpper();
                int anio = int.Parse(match.Groups[3].Value);

                if (mesesEspanol.TryGetValue(mesTexto, out int mes))
                {
                    // Construir la fecha
                    fecha = new DateTime(anio, mes, dia);
                    return true;
                }
            }

            return false;
        }
    }
}