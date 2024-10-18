namespace ProcesadorTxt
{
    public class MovimientoArticulo
    {
        public string Grupo { get; set; }
        public string Articulo { get; set; }
        public string Descripcion { get; set; }
        public string Uni { get; set; }
        public string CantidadPresentacion { get; set; }
        public string Tipo { get; set; }
        public string Puu { get; set; }
        public string ExistenciaInicial { get; set; }
        public string TipoMovimiento { get; set; }
        public string Documento { get; set; }
        public string UnidadProveedor { get; set; }
        public string Fecha { get; set; }
        public string Entradas { get; set; }
        public string Salidas { get; set; }
        public string Saldos { get; set; }
        public string NombreProveedor { get; set; }
        public string Lote { get; set; }
        public string Caducidad { get; set; }
        public string Cantidad { get; set; }
        public string ExistenciaFinalEntradas { get; set; } = "0"; // Se inicializan en "0"
        public string ExistenciaFinalSalidas { get; set; } = "0";
        public string ExistenciaFinalSaldos { get; set; } = "0";
    }
}