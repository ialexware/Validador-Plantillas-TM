using Newtonsoft.Json;

namespace Validador_Plantillas_TM.Modelos
{
    public class ClaveUnidad
    {
        [JsonProperty("c_ClaveUnidad")]
        public string CClaveUnidad { get; set; }

        [JsonProperty("Nombre")]
        public string Nombre { get; set; }

        [JsonProperty("Nota", NullValueHandling = NullValueHandling.Ignore)]
        public string Nota { get; set; }

        [JsonProperty("FechaDeInicioDeVigencia")]
        public string FechaDeInicioDeVigencia { get; set; }

        [JsonProperty("Símbolo", NullValueHandling = NullValueHandling.Ignore)]
        public string Símbolo { get; set; }

        [JsonProperty("Descripción", NullValueHandling = NullValueHandling.Ignore)]
        public string Descripción { get; set; }
    }
}
