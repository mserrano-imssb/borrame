using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ProcesadorTxt;

partial class IVProcessorForm
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;
    private DataGridView dataGridView;
    private Button btnCargarArchivo;
    private Button btnExportarExcel;
    private System.Windows.Forms.Label lblSumDisponible;
    private System.Windows.Forms.Label lblSumImporte;
    private System.Windows.Forms.TextBox txtSubtotales;
    private Button btnRegresar;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
            // Inicialización de componentes existentes
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.btnCargarArchivo = new System.Windows.Forms.Button();
            this.btnExportarExcel = new System.Windows.Forms.Button();
            this.txtSubtotales = new System.Windows.Forms.TextBox();
            this.lblSumDisponible = new System.Windows.Forms.Label();
            this.lblSumImporte = new System.Windows.Forms.Label();
            this.btnRegresar = new Button();
            this.dataTable = new System.Data.DataTable();

            // Configuración del DataGridView
            this.dataGridView.Location = new System.Drawing.Point(20, 20);
            this.dataGridView.Size = new System.Drawing.Size(800, 400);

            // Configuración del botón de cargar archivo
            this.btnCargarArchivo.Text = "Cargar Archivo";
            this.btnCargarArchivo.Location = new System.Drawing.Point(20, 440);
            this.btnCargarArchivo.Click += new System.EventHandler(this.BtnCargarArchivo_Click);

            // Configuración del botón de exportar a Excel
            this.btnExportarExcel.Text = "Exportar a Excel";
            this.btnExportarExcel.Location = new System.Drawing.Point(150, 440);
            this.btnExportarExcel.Click += new System.EventHandler(this.btnExportarExcel_Click);

            // Configuración del TextBox para mostrar los subtotales
            this.txtSubtotales.Location = new System.Drawing.Point(210, 440);
            this.txtSubtotales.Size = new System.Drawing.Size(610, 150);
            this.txtSubtotales.Multiline = true;
            this.txtSubtotales.ReadOnly = true;
            this.txtSubtotales.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;

            // Configuración del Label para la sumatoria de "DISPONIBLE"
            this.lblSumDisponible.AutoSize = true;
            this.lblSumDisponible.Location = new System.Drawing.Point(20, 480);
            this.lblSumDisponible.Text = "Total DISPONIBLE: 0";

            // Configuración del Label para la sumatoria de "IMPORTE"
            this.lblSumImporte.AutoSize = true;
            this.lblSumImporte.Location = new System.Drawing.Point(20, 510);
            this.lblSumImporte.Text = "Total IMPORTE: 0";

            // Configuración del botón de regresar
            this.btnRegresar.Text = "Regresar";
            this.btnRegresar.Location = new System.Drawing.Point(20, 540);
            this.btnRegresar.Click += new EventHandler(this.BtnRegresar_Click);

            // Añadir controles al formulario
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.btnCargarArchivo);
            this.Controls.Add(this.btnExportarExcel);
            this.Controls.Add(this.lblSumDisponible);
            this.Controls.Add(this.lblSumImporte);
            this.Controls.Add(this.txtSubtotales);
            this.Controls.Add(this.btnRegresar);

            // Otras configuraciones del formulario
            this.Text = "Procesador de Archivo TXT";
            this.ClientSize = new System.Drawing.Size(850, 600);

        // Otras configuraciones del formulario
        this.Text = "Procesador de Archivo TXT de Inventarios Valorizados";

        // Configurar las columnas del DataTable
        ConfigurarDataTable();
    }

    private void ConfigurarDataTable()
    {
        // Añadir columnas al DataTable
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

        // Asignar el DataTable al DataGridView
        dataGridView.DataSource = dataTable;
    }

    private void BtnCargarArchivo_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog
        {
            Filter = "Text Files|*.txt",
            Title = "Seleccionar archivo TXT"
        };

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            ProcesarArchivo(openFileDialog.FileName);
        }
    }

    #endregion
}
