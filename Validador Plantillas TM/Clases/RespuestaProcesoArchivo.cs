using System.Collections.Generic;

namespace Validador_Plantillas_TM.Clases
{
    public class RespuestaProcesoArchivo
    {
        public string strNombreHoja { get; set; }
        public int intTotalFilas { get; set; } = 0;
        public int intTotalHojas { get; set; } = 0;
        public int intTotalColumnas { get; set; } = 0;
        public int intTotalFilasOk { get; set; } = 0;
        public int intTotalFilasError { get; set; } = 0;
        public List<Incidencia> lstIncidencias { get; set; } = new List<Incidencia>();
    }
}
