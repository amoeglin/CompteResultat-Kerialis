using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompteResultat.DAL
{
    public class GenericClasses
    {
        public GenericClasses() { }

        public string GroupName { get; set; }
        public string GarantyName { get; set; }
        public string CodeActe { get; set; }
        public string AssureurName { get; set; }

    }

    public class GroupesGarantiesSante
    {
        public string AssureurName { get; set; }
        public string GroupName { get; set; }
        public string GarantyName { get; set; }
        public string CodeActe { get; set; }        
        public int OrderNumber { get; set; }
    }

    public class ExcelGlobalDecompteData
    {
        public string Assureur { get; set; }
        public string Company { get; set; }
        public string Subsid { get; set; }
        public int YearSurv { get; set; }
        public double? FR { get; set; }

        public double? RSS { get; set; }
        public double? RAnnexe { get; set; }
        public double? RNous { get; set; }
        public double Provisions { get; set; }
        public double CotBrute { get; set; }
        public string TaxTotal { get; set; }
        public string TaxDefault { get; set; }
        public string TaxActive { get; set; }
        public double CotNet { get; set; }
        public double Ratio { get; set; }
        public double GainLoss { get; set; }
        public DateTime DateArret { get; set; }
        public double? Coef { get; set; }
    }

    public class ExcelGlobalSinistreData
    {
        public string Assureur { get; set; }
        public string Company { get; set; }
        public string Subsid { get; set; }
        public int YearSurv { get; set; }
        public double? FR { get; set; }
        public double? RSS { get; set; }
        public double? RAnnexe { get; set; }
        public double? RNous { get; set; }
        public double Provisions { get; set; }
        public double CotBrut { get; set; }
        public string TaxTotal { get; set; }
        public string TaxDefault { get; set; }
        public string TaxActive { get; set; }
        public double CotNet { get; set; }
        public double Ratio { get; set; }
        public double GainLoss { get; set; }
        public DateTime DateArret { get; set; }

        public double? Coef { get; set; }
    }

    public class ExcelGlobalPrestaData
    {
        public string Assureur { get; set; }
        public string Company { get; set; }
        public string Subsid { get; set; }
        public int YearSurv { get; set; }       
        public double? FR { get; set; }
        public double? RSS { get; set; }
        public double? RAnnexe { get; set; }
        public double? RNous { get; set; }
        public double Provisions { get; set; }
        public double CotBrut { get; set; }
        public string TaxTotal { get; set; }
        public string TaxDefault { get; set; }
        public string TaxActive { get; set; }
        public double CotNet { get; set; }
        public double Ratio { get; set; }
        public double GainLoss { get; set; }
        public DateTime DateArret { get; set; }
      
        public double? Coef { get; set; }
    }

    public class ExcelGlobalCotisatData
    {
        public string Assureur { get; set; }
        public string Company { get; set; }
        public string Subsid { get; set; }
        public int YearSurv { get; set; }
        public double? Cotisat { get; set; }
        public double? CotisatBrute { get; set; }
    }
    public class IMAssurContrIDPair : IEquatable<IMAssurContrIDPair>
    {
        public IMAssurContrIDPair() { }

        public int IdAssurance { get; set; }
        public int IdContract { get; set; }

        public override string ToString()
        {
            return IdAssurance + "-" + IdContract;
        }

        //public bool Equals(IMAssurContrIDPair other)
        //{
        //    if (other == null) return false;
        //    return (this.IdAssurance.Equals(other.IdAssurance) && this.IdContract.Equals(other.IdContract));
        //}

        public bool Equals(IMAssurContrIDPair other)
        {
            if (this.IdAssurance == other.IdAssurance && this.IdContract == other.IdContract)            
                return true;            
            else           
                return false;            
        }

        //public override bool Equals(object obj)
        //{
        //    if (obj == null) return false;
        //    IMAssurContrIDPair objAsPart = obj as IMAssurContrIDPair;

        //    if (objAsPart == null) return false;
        //    else return Equals(objAsPart);
        //}

       

        //Using:
        //parts.Contains(new Part { PartId = 1734, PartName = "" });

        // Find items where name contains "seat".
        //parts.Find(x => x.PartName.Contains("seat"));

        // Check if an item with Id 1444 exists.
        //parts.Exists(x => x.PartId == 1444);

    }

    public class IMContrCompIDPair : IEquatable<IMContrCompIDPair>
    {
        public IMContrCompIDPair() { }
        
        public int IdContract { get; set; }
        public int IdCompany { get; set; }

        public override string ToString()
        {
            return IdContract + "-" + IdCompany;
        }

        public bool Equals(IMContrCompIDPair other)
        {
            if (this.IdContract == other.IdContract && this.IdCompany == other.IdCompany)
                return true;
            else
                return false;
        }
    }

    public class CompNameIDPair : IEquatable<CompNameIDPair>
    {
        public CompNameIDPair() { }

        public string CompanyName { get; set; }
        public int CompanyId { get; set; }

        public override string ToString()
        {
            return CompanyId.ToString() + "-" + CompanyName;
        }

        public bool Equals(CompNameIDPair other)
        {
            if (this.CompanyId == other.CompanyId && this.CompanyName == other.CompanyName)
                return true;
            else
                return false;
        }
    }

    public class ContrNameIDPair : IEquatable<ContrNameIDPair>
    {
        public ContrNameIDPair() { }

        public string ContrName { get; set; }
        public int ContrId { get; set; }

        public override string ToString()
        {
            return ContrId.ToString() + "-" + ContrName;
        }

        public bool Equals(ContrNameIDPair other)
        {
            if (this.ContrId == other.ContrId && this.ContrName == other.ContrName)
                return true;
            else
                return false;
        }
    }

    public class OtherTableAssurContrPair
    {
        public OtherTableAssurContrPair() { }

        public string Assureur { get; set; }
        public string ContractId { get; set; }
    }

    public class OtherTableContrCompPair
    {
        public OtherTableContrCompPair() { }
        
        public string ContractId { get; set; }
        public string Company { get; set; }

    }

    public class OtherTableContrSubsidPair
    {
        public OtherTableContrSubsidPair() { }

        public string ContractId { get; set; }
        public string Subsid { get; set; }

    }

    public class DecomptePrevReduced
    {
        public DecomptePrevReduced() { }

        public string Dossier { get; set; }
        public double? Total { get; set; }
        public DateTime? DatePayement { get; set; }
        public DateTime? DebSin { get; set; }
        public DateTime? FinSin { get; set; }
    }

    public class DemoSanteWithOptionInfo
    {
        public DemoSanteWithOptionInfo() { }

        public string ContractId { get; set; }
        public string WithOption { get; set; }
    }

    public class PrestSanteContrIdCount
    {
        public PrestSanteContrIdCount() { }

        public string ContractId { get; set; }
        public long Count { get; set; }
    }

}
