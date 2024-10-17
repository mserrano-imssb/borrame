using System;
using System.Windows.Forms;

namespace ProcesadorTxt
{
    public partial class MainForm : Form
    {
        private const string TituloPrincipal = "Procesador de TXT y otras utilerias";
        private Panel sidebarPanel;
        private Button btnHome;
        private Button btnSubMenu;
        private Button btnSettings;
        private Button btnArticulosProcessor;
        private Button BtnIVProcessor;
        private Panel contentPanel;
        private Button currentButton; // Botón actual seleccionado

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.sidebarPanel = new Panel();
            this.btnHome = new Button();
            this.btnSubMenu = new Button();
            this.btnSettings = new Button();
            this.btnArticulosProcessor = new Button();
            this.BtnIVProcessor = new Button();
            this.contentPanel = new Panel();

            // Configuración del Panel de la barra lateral
            this.sidebarPanel.BackColor = System.Drawing.Color.FromArgb(30, 30, 30);
            this.sidebarPanel.Dock = DockStyle.Left;
            this.sidebarPanel.Width = 200;
            this.sidebarPanel.Controls.Add(this.btnHome);
            this.sidebarPanel.Controls.Add(this.btnSubMenu);
            this.sidebarPanel.Controls.Add(this.btnSettings);
            this.sidebarPanel.Controls.Add(this.btnArticulosProcessor);
            this.sidebarPanel.Controls.Add(this.BtnIVProcessor);

           // Configuración de los botones de la barra lateral
            ConfigureButton(this.btnHome, "Home", new EventHandler(this.BtnHome_Click));
            ConfigureButton(this.btnSubMenu, "Sub Menu", new EventHandler(this.BtnSubMenu_Click));
            ConfigureButton(this.btnSettings, "Settings", new EventHandler(this.BtnSettings_Click));
            ConfigureButton(this.btnArticulosProcessor, "Cat. Artículos", new EventHandler(this.BtnArticulosProcessor_Click));
            ConfigureButton(this.BtnIVProcessor, "Inventario Valorizado", new EventHandler(this.BtnIVProcessor_Click));

            // Configuración del Panel de contenido
            this.contentPanel.Dock = DockStyle.Fill;
            this.contentPanel.BackColor = System.Drawing.Color.White;

            // Añadir los paneles al formulario principal
            this.Controls.Add(this.contentPanel);
            this.Controls.Add(this.sidebarPanel);

            // Configuración del formulario principal
            this.Text = TituloPrincipal;
            this.Size = new System.Drawing.Size(800, 600);
        }

        private void ConfigureButton(Button button, string text, EventHandler clickEvent)
        {
            button.Text = text;
            button.Dock = DockStyle.Top;
            button.FlatStyle = FlatStyle.Flat;
            button.ForeColor = System.Drawing.Color.White;
            button.BackColor = System.Drawing.Color.FromArgb(30, 30, 30); // Color normal
            button.Height = 50;
            button.Click += clickEvent;
        }

        private void SetActiveButton(Button button)
        {
            if (currentButton != null)
            {
                currentButton.BackColor = System.Drawing.Color.FromArgb(30, 30, 30); // Restaurar el color del botón anterior
            }

            // Cambiar el color del botón seleccionado
            currentButton = button;
            currentButton.BackColor = System.Drawing.Color.FromArgb(90, 90, 90); // Nuevo color para el botón activo
        }

        private void BtnHome_Click(object sender, EventArgs e)
        {
            SetActiveButton((Button)sender);
            // Cambiar el contenido principal cuando se presiona el botón "Home"
            this.contentPanel.Controls.Clear();
            Label lbl = new Label();
            lbl.Text = "Home Content";
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.contentPanel.Controls.Add(lbl);
        }

        private void BtnSubMenu_Click(object sender, EventArgs e)
        {
            SetActiveButton((Button)sender);
            // Cambiar el contenido principal cuando se presiona el botón "Sub Menu"
            this.contentPanel.Controls.Clear();
            Label lbl = new Label();
            lbl.Text = "Sub Menu Content";
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.contentPanel.Controls.Add(lbl);
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            SetActiveButton((Button)sender);
            // Cambiar el contenido principal cuando se presiona el botón "Settings"
            this.contentPanel.Controls.Clear();
            Label lbl = new Label();
            lbl.Text = "Settings Content";
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.contentPanel.Controls.Add(lbl);
        }

        private void BtnArticulosProcessor_Click(object sender, EventArgs e)
        {
            SetActiveButton((Button)sender);
            // Cambiar el contenido principal cuando se presiona el botón "Help"
            this.contentPanel.Controls.Clear();
           /* Label lbl = new Label();
            lbl.Text = "Help Content";
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.contentPanel.Controls.Add(lbl);*/
            // Crear una instancia del formulario ArticulosProcessorForm
            ArticulosProcessorForm articulosForm = new ArticulosProcessorForm();

            // Configurar el formulario para que se comporte como un control secundario del panel
            articulosForm.TopLevel = false;  // Esto es importante para que el formulario no se muestre como una ventana independiente
            articulosForm.FormBorderStyle = FormBorderStyle.None;
            articulosForm.Dock = DockStyle.Fill;  // Esto hace que el formulario ocupe todo el espacio del contentPanel

            // Agregar el formulario al contentPanel
            this.contentPanel.Controls.Add(articulosForm);
            articulosForm.Show();  // Mostrar el formulario
            this.Text = TituloPrincipal + " (TXT de Articulos)";
        }

        private void BtnIVProcessor_Click(object sender, EventArgs e)
        {
            SetActiveButton((Button)sender);
            // Cambiar el contenido principal cuando se presiona el botón "About"
            this.contentPanel.Controls.Clear();
            /*Label lbl = new Label();
            lbl.Text = "About Content";
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.contentPanel.Controls.Add(lbl);*/
             // Crear una instancia del formulario IVProcessorForm
            IVProcessorForm ivForm = new IVProcessorForm();

            // Configurar el formulario para que se comporte como un control secundario del panel
            ivForm.TopLevel = false;  // Esto es importante para que el formulario no se muestre como una ventana independiente
            ivForm.FormBorderStyle = FormBorderStyle.None;
            ivForm.Dock = DockStyle.Fill;  // Esto hace que el formulario ocupe todo el espacio del contentPanel

            // Agregar el formulario al contentPanel
            this.contentPanel.Controls.Add(ivForm);
            ivForm.Show();  // Mostrar el formulario
            this.Text = TituloPrincipal + " (TXT de Inventario Valorizado)";
        }
    }
}
