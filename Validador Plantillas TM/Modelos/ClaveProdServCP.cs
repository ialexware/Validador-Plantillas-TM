using Newtonsoft.Json;

namespace Validador_Plantillas_TM.Modelos
{
    public class ClaveProdServCP
    {
        [JsonProperty("c_ClaveProdServ")]
        public string CClaveProdServ { get; set; }

        [JsonProperty("Descripción")]
        public string Descripción { get; set; }

        [JsonProperty("Palabras similares")]
        public string PalabrasSimilares { get; set; }

        [JsonProperty("Material Peligroso")]
        public string MaterialPeligroso { get; set; }

        public bool blnMaterialPeligroso
        {
            get
            {
                switch (MaterialPeligroso)
                {
                    case "0":
                        return false;
                    case "1":
                        return true;
                    default:
                        return false;
                }
            }
        }

        [JsonProperty("FechaInicioVigencia")]
        public string FechaInicioVigencia { get; set; }

        [JsonProperty("FechaFinVigencia")]
        public object FechaFinVigencia { get; set; }
    }
}
