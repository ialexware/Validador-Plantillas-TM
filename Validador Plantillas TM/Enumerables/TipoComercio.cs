namespace Validador_Plantillas_TM.Enumerables
{
    public enum TipoComercioPlantilla
    {
        Importacion,
        Exportacion,
        Local
    }

    public static class TipoComercio
    {
        public static string GetStringTipoComercio(this TipoComercioPlantilla objTipo)
        {
            switch (objTipo)
            {
                case TipoComercioPlantilla.Importacion:
                    return "Importación";
                case TipoComercioPlantilla.Exportacion:
                    return "Exportación";
                case TipoComercioPlantilla.Local:
                    return string.Empty;
                default:
                    return string.Empty;
            }
        }
    }
}
