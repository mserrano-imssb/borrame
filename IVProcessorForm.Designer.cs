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
    private TableLayoutPanel mainLayoutPanel;
    private FlowLayoutPanel buttonPanel;

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
        // Inicialización del TableLayoutPanel principal
        this.mainLayoutPanel = new TableLayoutPanel();
        this.mainLayoutPanel.ColumnCount = 1;
        this.mainLayoutPanel.RowCount = 4;
        this.mainLayoutPanel.Dock = DockStyle.Fill;
        this.mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 120F)); // 120% para el DataGridView
        this.mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F)); // Fila para botones
        this.mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 60F)); // Fila para subtotales
        this.mainLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 5F)); // Fila para labels sumatorias

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
        this.dataGridView.Dock = DockStyle.Fill;        
        this.dataGridView.Size = new System.Drawing.Size(800, 400);
        //this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        // Configuración del botón de cargar archivo
        this.btnCargarArchivo.Text = "Cargar Archivo";
        this.btnCargarArchivo.Size = new System.Drawing.Size(120, 30);
        this.btnCargarArchivo.Click += new System.EventHandler(this.BtnCargarArchivo_Click);

        // Configuración del botón de exportar a Excel
        this.btnExportarExcel.Text = "Exportar a Excel";
        this.btnExportarExcel.Size = new System.Drawing.Size(120, 30);
        this.btnExportarExcel.Click += new System.EventHandler(this.btnExportarExcel_Click);

        // Configuración del TextBox para mostrar los subtotales
        this.txtSubtotales.Dock = DockStyle.Fill;
        this.txtSubtotales.Multiline = true;
        this.txtSubtotales.ReadOnly = true;
        this.txtSubtotales.ScrollBars = ScrollBars.Vertical;

        // Configuración del Label para la sumatoria de "DISPONIBLE"
        this.lblSumDisponible.AutoSize = true;
        this.lblSumDisponible.Dock = DockStyle.Top;
        this.lblSumDisponible.Text = "Total DISPONIBLE: 0";

        // Configuración del Label para la sumatoria de "IMPORTE"
        this.lblSumImporte.AutoSize = true;
        this.lblSumImporte.Dock = DockStyle.Top;
        this.lblSumImporte.Text = "Total IMPORTE: 0";

        // Configuración del botón de regresar
        this.btnRegresar.Text = "Regresar";
        this.btnRegresar.Size = new System.Drawing.Size(120, 30);
        this.btnRegresar.Click += new EventHandler(this.BtnRegresar_Click);

        // Panel para los botones
        buttonPanel = new FlowLayoutPanel();
        buttonPanel.Dock = DockStyle.Fill;
        buttonPanel.FlowDirection = FlowDirection.LeftToRight; // Botones de izquierda a derecha
        buttonPanel.WrapContents = false;  // Evitar que los botones se apilen en múltiples líneas
        buttonPanel.AutoSize = true; // Permitir que el panel se ajuste al contenido

        // Añadir los botones al buttonPanel
        buttonPanel.Controls.Add(this.btnCargarArchivo);
        buttonPanel.Controls.Add(this.btnExportarExcel);
       // buttonPanel.Controls.Add(this.btnRegresar); // temporalmente

        // Agregar controles al TableLayoutPanel
        this.mainLayoutPanel.Controls.Add(this.dataGridView, 0, 0);  // Primera fila (DataGridView)
        this.mainLayoutPanel.Controls.Add(buttonPanel, 0, 1);       // Segunda fila (Botones)
        this.mainLayoutPanel.Controls.Add(this.txtSubtotales, 0, 2); // Tercera fila (Subtotales)
        this.mainLayoutPanel.Controls.Add(this.lblSumDisponible, 0, 3); // Cuarta fila (Sumatoria de disponible)
        this.mainLayoutPanel.Controls.Add(this.lblSumImporte, 0, 3);    // Cuarta fila (Sumatoria de importe)

        // Añadir el TableLayoutPanel al formulario
        this.Controls.Add(this.mainLayoutPanel);

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
