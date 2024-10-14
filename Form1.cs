using System;
using System.Windows.Forms;
using ProcesadorTxt;

namespace ProcesadorTxt
{
    public partial class Form1 : Form
    {
        private Button btnIVProcessor;
        private ToolTip toolTipIV;

        private Button btnArticulosProcessor;
        private ToolTip toolTipArticulos;

        public Form1()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.btnIVProcessor = new();
            this.toolTipIV = new();
            this.btnArticulosProcessor = new();
            this.toolTipArticulos = new();

            // Configuración del botón IVProcessor
            this.btnIVProcessor.Text = "IV Processor";
            this.btnIVProcessor.Location = new System.Drawing.Point(20, 20);
            this.btnIVProcessor.Click += new EventHandler(this.BtnIVProcessor_Click);
            
            // Configuración del botón ArticulosProcessor
            this.btnArticulosProcessor.Text = "Articulos Processor";
            this.btnArticulosProcessor.Location = new System.Drawing.Point(20, 80);
            this.btnArticulosProcessor.Click += new EventHandler(this.BtnArticulosProcessor_Click);

            // Configuración del ToolTip
            toolTipIV.SetToolTip(this.btnIVProcessor, "Clic para abrir el procesador de txt de Inventario Valorizado");  // Mensaje del ToolTip
            toolTipArticulos.SetToolTip(this.btnArticulosProcessor, "Clic para abrir el procesador de txt de Articulos");            

            // Añadir el botón al formulario
            this.Controls.Add(this.btnIVProcessor);
            this.Controls.Add(this.btnArticulosProcessor);

            // Configuración del formulario
            this.Text = "IMSSB - Mini utilerías";
            this.Size = new System.Drawing.Size(400, 300);
        }

        private void BtnIVProcessor_Click(object sender, EventArgs e)
        {
            // Mostrar el formulario IVProcessorForm
            IVProcessorForm ivForm = new IVProcessorForm();
            this.Hide();  // Esconde el formulario principal
            ivForm.Show();
        }

        private void BtnArticulosProcessor_Click(object sender, EventArgs e)
        {
            // Mostrar el formulario IVProcessorForm
            ArticulosProcessorForm articulosForm = new ArticulosProcessorForm();
            this.Hide();  // Esconde el formulario principal
            articulosForm.Show();
        }
    }
}
