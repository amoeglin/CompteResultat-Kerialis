using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using CompteResultat.Common;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity.Validation;

namespace CompteResultat.DAL
{
    [MetadataType(typeof(PrestSante.MetaData))]
    public partial class PrestSante
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static List<ExcelGlobalPrestaData> GetPrestaGlobalEntData(List<int> years, List<string> companyList)
        {
            try
            {
                List<ExcelGlobalPrestaData> prestations = new List<ExcelGlobalPrestaData>();

                using (var context = new CompteResultatEntities())
                {
                    prestations = context.PrestSantes
                    .Where(d => years.Contains(d.DateSoins.Value.Year) && companyList.Contains(d.Company))
                    .GroupBy(p => new { p.AssureurName, p.Company, AnnSurv = p.DateSoins.Value.Year })
                    .Select(g => new ExcelGlobalPrestaData
                    {
                        Assureur = g.Key.AssureurName,
                        Company = g.Key.Company,
                        Subsid = "",
                        YearSurv = g.Key.AnnSurv,
                        FR = g.Sum(i => i.FraisReel),
                        RSS = g.Sum(i => i.RembSS),
                        RAnnexe = g.Sum(i => i.RembAnnexe),
                        RNous = g.Sum(i => i.RembNous),
                        Provisions = 0,
                        CotBrut = 0,
                        TaxTotal = "",
                        TaxDefault = "",
                        TaxActive = "",
                        CotNet = 0,
                        Ratio = 0,
                        GainLoss = 0,
                        DateArret = DateTime.Now
                    })
                    //.OrderBy(ga => ga.YearSurv).ThenBy(gb => gb.Company)
                    .OrderBy(ga => ga.Company).ThenBy(gb => gb.Subsid).ThenBy(gc => gc.YearSurv)
                    .ToList();
                }

                return prestations;

            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static List<ExcelGlobalPrestaData> GetPrestaGlobalSubsidData(List<int> years, List<string> subsidList)
        {
            try
            {
                List<ExcelGlobalPrestaData> prestations = new List<ExcelGlobalPrestaData>();

                using (var context = new CompteResultatEntities())
                {
                    prestations = context.PrestSantes
                    .Where(d => years.Contains(d.DateSoins.Value.Year) && subsidList.Contains(d.Company))
                    .GroupBy(p => new { p.AssureurName, p.Company, p.Subsid, AnnSurv = p.DateSoins.Value.Year })
                    .Select(g => new ExcelGlobalPrestaData
                    {
                        Assureur = g.Key.AssureurName,
                        Company = g.Key.Company,
                        Subsid = g.Key.Subsid,
                        YearSurv = g.Key.AnnSurv,
                        FR = g.Sum(i => i.FraisReel),
                        RSS = g.Sum(i => i.RembSS),
                        RAnnexe = g.Sum(i => i.RembAnnexe),
                        RNous = g.Sum(i => i.RembNous),
                        Provisions = 0,
                        CotBrut = 0,
                        TaxTotal = "",
                        TaxDefault = "",
                        TaxActive = "",
                        CotNet = 0,
                        Ratio = 0,
                        GainLoss = 0,
                        DateArret = DateTime.Now
                    })
                    //.OrderBy(ga => ga.YearSurv).ThenBy(gb => gb.Company).ThenBy(gc => gc.Subsid)
                    .OrderBy(ga => ga.Company).ThenBy(gb => gb.Subsid).ThenBy(gc => gc.YearSurv)
                    .ToList();
                }

                return prestations;

            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static List<PrestSante> GetPrestationsForContracts(List<string> assurList, List<string> parentCompanyList, List<string> companyList, 
            List<string> contrIds, string college, DateTime debutPeriod, DateTime finPeriod, DateTime dateArret)
        {
            try
            {
                List<PrestSante> prestations;

                using (var context = new CompteResultatEntities())
                {
                    prestations = context.PrestSantes.Where(prest => assurList.Contains(prest.AssureurName) && parentCompanyList.Contains(prest.Company)
                        && companyList.Contains(prest.Subsid) && contrIds.Contains(prest.ContractId)
                        && prest.DateSoins >= debutPeriod && prest.DateSoins <= finPeriod && prest.DatePayment <= dateArret).ToList();

                    //prestations = context.PrestSantes.Where(prest => contrIds.Contains(prest.ContractId) 
                    //   && prest.DateSoins >= debutPeriod && prest.DateSoins <= finPeriod && prest.DatePayment <= dateArret).ToList();

                    //var x = prestations
                    //       .Where(prest => contrIds.Contains(prest.ContractId) && prest.DateSoins >= debutPeriod && prest.DateSoins <= finPeriod)
                    //       .GroupBy(p => new { p.DateVision, p.ContractId, p.CodeCol, DateSoinsYear = p.DateSoins.Value.Year, p.CodeActe, p.CAS, p.Reseau })
                    //       .Select(p => new
                    //       {
                    //           DateVision = p.Key.DateVision,
                    //           ContractId = p.Key.ContractId,
                    //           CodeCol = p.Key.CodeCol,
                    //           DateSoins = new DateTime(p.Key.DateSoinsYear, 1, 1),
                    //           CodeActe = p.Key.CodeActe,
                    //           CAS = p.Key.CAS,
                    //           NombreActe = p.Sum(pr => pr.NombreActe),
                    //           FraisReel = p.Sum(pr => pr.FraisReel),
                    //           RembSS = p.Sum(pr => pr.RembSS),
                    //           RembAnnexe = p.Sum(pr => pr.RembAnnexe),
                    //           RembNous = p.Sum(pr => pr.RembNous),
                    //           Reseau = p.Key.Reseau
                    //       })
                    //       .ToList();


                    //taking into account College:
                    //prestations = context.PrestSantes.Where(prest => contrIds.Contains(prest.ContractId) && prest.CodeCol == college
                    //    && prest.DateSoins >= debutPeriod && prest.DateSoins <= finPeriod).ToList();

                    //cotisat2 = context.CotisatSantes.Where(cot => contrIds.Contains(cot.ContractId)).
                    //    Select(cot => new { cot.DebPrime, cot.FinPrime, cot.ContractId, cot.CodeCol, cot.Year, cot.Cotisation }).ToList();
                }

                return prestations;

            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static List<PrestSanteContrIdCount> GetContractIdCount()
        {
            try
            {
                List<PrestSanteContrIdCount> contrCount;

                using (var context = new CompteResultatEntities())
                {
                    contrCount = context.PrestSantes
                        .GroupBy(p => new { p.ContractId })
                        .Select(p => new PrestSanteContrIdCount
                        {
                            ContractId = p.Key.ContractId,
                            Count = p.Count()
                        })
                        .OrderBy(gr => gr.Count)
                        .ToList();
                }

                return contrCount;

            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }
        
        public static List<GroupesGarantiesSante> GetGroupGarantyList()
        {
            try
            {
                List<GroupesGarantiesSante> GroupGarantyTable;

                using (var context = new CompteResultatEntities())
                {
                    GroupGarantyTable = context.Database
                            .SqlQuery<GroupesGarantiesSante>("SELECT DISTINCT AssureurName,GroupName,GarantyName,CodeActe,OrderNumber=1 FROM dbo.PrestSante ORDER BY AssureurName,GroupName,GarantyName,CodeActe")
                            .ToList<GroupesGarantiesSante>();
                }

                return GroupGarantyTable;
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }
        
        public static void DeleteRowsWithImportId(int importId)
        {
            try
            {
                using (var context = new CompteResultatEntities())
                {
                    //context.PrestSantes.RemoveRange(context.PrestSantes.Where(c => c.ImportId == importId));
                    //context.SaveChanges();

                    context.Database.ExecuteSqlCommand("DELETE FROM PrestSante WHERE ImportId = {0}", importId);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static void UpdateOptionField(int importId, string contractId)
        {
            try
            {
                
                using (var context = new CompteResultatEntities())
                {

                    context.Database.ExecuteSqlCommand("update PrestSante set WithOption = 'True' where ContractId = '" + contractId + "' and ImportId = " + importId);

                    //var elements = context.PrestSantes.Where(p => p.ContractId == contractId && p.ImportId == importId);

                    //if (elements.Any())
                    //{
                    //    PrestSante prest = elements.First();
                    //    prest.WithOption = "True";
                    //    context.SaveChanges();
                    //}
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static void UpdateWithOptionField(string sql)
        {
            try
            {

                using (var context = new CompteResultatEntities())
                {
                    context.Database.ExecuteSqlCommand(sql);
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }


        //MetaData definition for basic validation
        public class MetaData
        {
            //[Display(Name = "Email address")]
            //[Required(ErrorMessage = "The email address is required")]
            //public string Email { get; set; }

        }
    }
}
