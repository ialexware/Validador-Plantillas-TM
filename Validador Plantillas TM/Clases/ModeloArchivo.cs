using Validador_Plantillas_TM.Modelos;

namespace Validador_Plantillas_TM.Clases
{
    public class ModeloArchivo
    {
        public int intFila { get; set; }
        public string Origen { get; set; }
        public string DestinoFinal { get; set; }
        public string BienesTransp { get; set; }
        public ClaveProdServCP claveProdServCP { get; set; }
        public string Descripcion { get; set; }
        public string Cantidad { get; set; }
        public string ClaveUnidad { get; set; }
        public ClaveUnidad claveUnidad { get; set; }
        public string Unidad { get; set; }
        public string MaterialPeligroso { get; set; }
        public bool blnMaterialPeligroso { get; set; } = false;
        public string CveMaterialPeligroso { get; set; }
        public string Embalaje { get; set; }
        public string DescripEmbalaje { get; set; }
        public string PesoEnKg { get; set; }
        public string FraccionArancelaria { get; set; }
        public string UUIDComercioExt { get; set; }
        public string Pedimento { get; set; }
        public string Repartos { get; set; }
        public string Guia { get; set; }
    }
}
