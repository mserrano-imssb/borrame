using System.Text;
using ProcesadorTxt;

namespace ProcesadorTxt;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // Registrar proveedor de codificación para IBM437
       /* Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);*/
        
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }    
}