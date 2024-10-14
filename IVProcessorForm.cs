using OfficeOpenXml;
using OfficeOpenXml.Style;
using ProcesadorTxt;
using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;

namespace ProcesadorTxt;

public partial class IVProcessorForm : Form
{
    private DataTable dataTable;
    private string[] exclude = ["PZA", "ENV", "EQP", "AMP", "CJA", "JGO", "LTA", "BTE", "F.G", "FCO"];

    public IVProcessorForm()
    {
        InitializeComponent();
        dataTable = new DataTable();
        ConfigurarDataGridView();
    }

    private void ConfigurarDataGridView()
    {
        // Configuración de las columnas
        dataTable.Columns.Add("GRUPO", typeof(string));
        dataTable.Columns.Add("PARTIDA", typeof(string));
        dataTable.Columns.Add("GEN", typeof(string));
        dataTable.Columns.Add("ESP", typeof(string));
        dataTable.Columns.Add("DIF", typeof(string));
        dataTable.Columns.Add("VAR", typeof(string));
        dataTable.Columns.Add("CLAVE", typeof(string));
        dataTable.Columns.Add("DESCRIPCION", typeof(string));
        dataTable.Columns.Add("UNI", typeof(string));
        dataTable.Columns.Add("CANT.P", typeof(decimal));
        dataTable.Columns.Add("TIPO", typeof(string));
        dataTable.Columns.Add("P.U.U.", typeof(decimal));
        dataTable.Columns.Add("CPM_V", typeof(int));
        dataTable.Columns.Add("P/EMBARQUE", typeof(int));
        dataTable.Columns.Add("EN EMBARQUE", typeof(int));
        dataTable.Columns.Add("DISPONIBLE", typeof(int));
        dataTable.Columns.Add("NI(DISP)", typeof(int));
        dataTable.Columns.Add("IMPORTE", typeof(decimal));

        dataGridView.DataSource = dataTable;
    }

    private void btnCargarArchivo_Click(object sender, EventArgs e)
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

    private void ProcesarArchivo(string filePath)
    {
        var line = "";
        try
        {
            // Registrar proveedor de codificación para páginas de códigos, necesario para codificaciones como Windows-1252
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Leer el archivo con la codificación ANSI (Windows-1252)
            string[] lines = File.ReadAllLines(filePath, Encoding.GetEncoding("Windows-1252"));
            string grupo = "";
            string grupoAnterior = "";
            string partida = "";
            decimal subtotalDisponible = 0;
            decimal subtotalImporte = 0;

            decimal totalDisponible = 0;
            decimal totalImporte = 0;
            this.dataTable.Rows.Clear();

            foreach (string linea in lines)
            {
                line = linea.Trim();
                // Detectar grupo y partida
                if (line.Contains("GRUPO") && !line.Contains("TOT.GRUPO :"))
                {
                    /* obtener el trozo de la cadena empezando con GRUPO, dado que puede haber antes un trozo que tambien contiene ':'
                    ejemplo "EXISTENCIAS AL CORTE DE : 2024/09/25                                                  GRUPO : 010/MEDICINAS."
                    */
                    line = line.Replace("EXISTENCIAS AL CORTE DE : ", "");
                    grupo = line.Split(':')[1].Trim().Split('/')[0];
                    if (grupoAnterior != grupo)
                    {
                        // Si estamos cambiando de grupo, muestra subtotales del grupo anterior
                        if (!string.IsNullOrEmpty(grupoAnterior))
                        {
                            MostrarSubtotales(grupoAnterior, partida, subtotalDisponible, subtotalImporte);
                            subtotalDisponible = 0;
                            subtotalImporte = 0;
                        }
                        grupoAnterior = grupo;
                    }
                }
                else if (line.Contains("PARTIDA"))
                {
                    partida = line.Split(':')[1].Trim().Split('/')[0];
                } // Detectar líneas de datos
                else if (line.Length > 0 && char.IsDigit(line[0]))  // Las líneas que empiezan con un número contienen los datos
                {
                    string[] preColumns = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    // Evitar errores si el formato no es correcto
                    if (preColumns.Length < 14) continue;

                    /* Desde la preColumns[4] voy acumulando los datos de las demas hasta que su longitud sea menor o igual a 30 */
                    List<string> columnsList = new List<string>();
                    var descripcion = "";
                    columnsList.Add(preColumns[0]);
                    columnsList.Add(preColumns[1]);
                    /*if (preColumns[0] == "711" && preColumns[1] == "0145")
                    {
                        // quiero solo mostrar un modal de OK o un dialogo de confirmacion
                        var popo = 0;
                    }*/
                    columnsList.Add(preColumns[2]);
                    columnsList.Add(preColumns[3]);

                    string clave = preColumns[0] + "." + preColumns[1] + "." + preColumns[2] + "." + preColumns[3];
                    columnsList.Add(clave);

                    int i = 4;
                    for (i = 4; i < preColumns.Length &&
                                    descripcion.Length + preColumns[i].Length <= 30 &&
                                    !exclude.Contains(preColumns[i]);
                         i++)
                    {
                        descripcion += preColumns[i] + " ";
                    }
                    columnsList.Add(descripcion);
                    // agregar el resto de columnas
                    for (int j = i; j < preColumns.Length; j++)
                    {
                        columnsList.Add(preColumns[j]);
                    }

                    var columns = columnsList.ToArray(); // siempre debe ser 17
                                                         // Parsear datos de las columnas "DISPONIBLE" y "IMPORTE"
                    int disponible = int.Parse(columns[13].Replace(",", ""));
                    decimal importe = decimal.Parse(columns[15].Replace("$", "").Replace(",", ""));

                    // Sumar subtotales y totales
                    subtotalDisponible += disponible;
                    subtotalImporte += importe;

                    // Sumar a los totales
                    totalDisponible += disponible;
                    totalImporte += importe;

                    dataTable.Rows.Add(
                        grupo,
                        partida,
                        columns[0],  // GEN
                        columns[1],  // ESP
                        columns[2],  // DIF
                        columns[3],  // VAR
                        columns[4],  // CLAVE
                        columns[5],  // DESCRIPCION
                        columns[6],  // UNI
                        decimal.Parse(columns[7].Replace(",", "")),  // CANT.P
                        columns[8],  // TIPO
                        decimal.Parse(columns[9].Replace("$", "").Replace(",", "")),  // P.U.U.
                        int.Parse(columns[10].Replace(",", "")),  // CPM_V
                        int.Parse(columns[11].Replace(",", "")),  // P/EMBARQUE
                        int.Parse(columns[12].Replace(",", "")),  // EN EMBARQUE
                        int.Parse(columns[13].Replace(",", "")),  // DISPONIBLE
                        int.Parse(columns[14].Replace(",", "")),  // NI(DISP)
                        decimal.Parse(columns[15].Replace("$", "").Replace(",", ""))  // IMPORTE
                    );
                }

                // Actualizar los Labels con las sumatorias
                var disponibleFormateado = totalDisponible.ToString("N");
                lblSumDisponible.Text = $"Total DISPONIBLE: {disponibleFormateado}";
                var importeFormateado = totalImporte.ToString("C3");
                lblSumImporte.Text = $"Total IMPORTE: {importeFormateado}";
            }
            // Mostrar subtotales finales para el último grupo/partida
            if (!string.IsNullOrEmpty(grupo))
            {
                MostrarSubtotales(grupo, partida, subtotalDisponible, subtotalImporte);
            }
        }
        catch (Exception ex)
        {
            throw;
        }
    }

    // Método para mostrar los subtotales en el TextBox
    private void MostrarSubtotales(string grupo, string partida, decimal disponible, decimal importe)
    {
        // Mostrar subtotales Importe en formato $xxx,xxx.xxx con 3 decimales
        var disponibleFormateado = disponible.ToString("N");
        var importeFormateado = importe.ToString("C3");
        txtSubtotales.AppendText($"Grupo: {grupo} - Partida: {partida} - Disponible: {disponibleFormateado} - Importe: {importeFormateado}\r\n");
    }

    private void btnExportarExcel_Click(object sender, EventArgs e)
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
        dataTable.DefaultView.Sort = "GRUPO, PARTIDA, GEN, ESP";
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

            /* package.SaveAs(new FileInfo(filePath));
             MessageBox.Show("Datos exportados con éxito a Excel", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);*/
            ExportarConTemplateAExcel(filePath);
            // ExportarConTemplateAExcelConFormulas(filePath);
        }
    }

    private void ExportarConTemplateAExcel(string filePath)
    {
        // Crear un nuevo archivo Excel
        using (var package = new ExcelPackage())
        {
            // Añadir una hoja llamada "Inventario"
            var worksheet = package.Workbook.Worksheets.Add("Inventario");

            // Crear el encabezado con el estilo deseado
            string[] header = { "GRUPO", "PARTIDA", "GEN", "ESP", "DIF", "VAR", "CLAVE",
                                 "DESCRIPCION", "UNI", "CANT.P", "TIPO", "P.U.U.", "CPM_V",
                                 "P/EMBARQUE", "EN EMBARQUE", "DISPONIBLE", "NI(DISP)", "IMPORTE" };

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

            MessageBox.Show("Archivo exportado con éxito con el formato de template.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>
    /// En construccion. No está 100 % correcto
    /// </summary>
    /// <param name="filePath"></param>
    private void ExportarConTemplateAExcelConFormulas(string filePath)
    {
        // Crear un nuevo archivo Excel
        using (var package = new ExcelPackage())
        {
            // Añadir una hoja llamada "Inventario"
            var worksheet = package.Workbook.Worksheets.Add("Inventario");

            // Crear el encabezado con el estilo deseado
            string[] header = { "GRUPO", "PARTIDA", "GEN", "ESP", "DIF", "VAR", "DESCRIPCION", "UNI", "CANT.P", "TIPO", "P.U.U.", "CPM_V", "P/EMBARQUE", "EN EMBARQUE", "DISPONIBLE", "NI(DISP)", "IMPORTE" };

            for (int i = 0; i < header.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = header[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                worksheet.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                worksheet.Cells[1, i + 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }

            // Poblar los datos en las filas debajo del encabezado
            int dataRowStart = 2;
            int dataRowEnd = dataRowStart + dataTable.Rows.Count - 1;

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    worksheet.Cells[i + dataRowStart, j + 1].Value = dataTable.Rows[i][j];
                }
            }

            // Autoajustar las columnas
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Insertar fórmula de total al final de la columna "DISPONIBLE" e "IMPORTE"
            int totalRow = dataRowEnd + 1;

            worksheet.Cells[totalRow, 14].Formula = $"SUM(N{dataRowStart}:N{dataRowEnd})";  // Sumar columna DISPONIBLE
            worksheet.Cells[totalRow, 14].Style.Font.Bold = true;

            worksheet.Cells[totalRow, 16].Formula = $"SUM(P{dataRowStart}:P{dataRowEnd})";  // Sumar columna IMPORTE
            worksheet.Cells[totalRow, 16].Style.Font.Bold = true;

            // Subtotales por grupo y partida
            string grupoAnterior = dataTable.Rows[0]["GRUPO"].ToString();
            string partidaAnterior = dataTable.Rows[0]["PARTIDA"].ToString();
            int subtotalStartRow = dataRowStart;

            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                string grupoActual = dataTable.Rows[i]["GRUPO"].ToString();
                string partidaActual = dataTable.Rows[i]["PARTIDA"].ToString();

                if (grupoActual != grupoAnterior || partidaActual != partidaAnterior)
                {
                    // Agregar subtotales debajo del bloque de grupo y partida anterior
                    int subtotalRow = subtotalStartRow + i;

                    worksheet.Cells[subtotalRow, 14].Formula = $"SUBTOTAL(9,N{subtotalStartRow}:N{subtotalRow - 1})";  // Subtotal DISPONIBLE
                    worksheet.Cells[subtotalRow, 14].Style.Font.Italic = true;

                    worksheet.Cells[subtotalRow, 16].Formula = $"SUBTOTAL(9,P{subtotalStartRow}:P{subtotalRow - 1})";  // Subtotal IMPORTE
                    worksheet.Cells[subtotalRow, 16].Style.Font.Italic = true;

                    // Actualizar el inicio del próximo subtotal
                    subtotalStartRow = subtotalRow + 1;

                    grupoAnterior = grupoActual;
                    partidaAnterior = partidaActual;
                }
            }

            // Guardar el archivo Excel en la ruta proporcionada
            FileInfo fi = new FileInfo(filePath);
            package.SaveAs(fi);

            MessageBox.Show("Archivo exportado con éxito con el formato de template y fórmulas.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
    private void BtnRegresar_Click(object sender, EventArgs e)
    {
        // Regresar al formulario principal
        this.Hide();
        Form1 mainForm = new Form1();
        mainForm.Show();
    }
}
