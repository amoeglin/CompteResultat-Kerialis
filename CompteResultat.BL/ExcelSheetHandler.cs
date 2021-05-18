using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity.Migrations;
using System.Data.Entity.Validation;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;
using System.Data;
using ServiceStack.Text;

using CompteResultat.DAL;
using CompteResultat.Common;

namespace CompteResultat.BL
{
    public class ExcelSheetHandler
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public string ReportTemplate { get; set; }

        public static void FillGlobalSheet(FileInfo excelFilePath, string companyList, string subsidList, DateTime debutPeriod,
           DateTime finPeriod, DateTime dateArret, C.eReportTypes reportType, double TaxDef, double TaxAct, double TaxPer)
        {
            try
            {
                bool isGlobalEnt = reportType == C.eReportTypes.GlobalEnt ? true : false;
                List<string> parentCompanyList = Regex.Split(companyList, C.cVALSEP).ToList();
                List<string> subsidiaryList = Regex.Split(subsidList, C.cVALSEP).ToList();
                DataTable globalTable = CreateGlobalTable(reportType);
                DataTable globalTableCumul = CreateGlobalTable(reportType);
                

                //get years && coefCad                
                List<int> years = new List<int>();
                for (int i = 0; i <= finPeriod.Year - debutPeriod.Year; i++)
                {
                    years.Add(debutPeriod.Year + i);
                }


                //get presta data
                List<ExcelGlobalPrestaData> globalPresta = new List<ExcelGlobalPrestaData>();
                if (isGlobalEnt)
                    globalPresta = PrestSante.GetPrestaGlobalEntData(years, parentCompanyList);
                else
                    globalPresta = PrestSante.GetPrestaGlobalSubsidData(years, subsidiaryList);

                //get cotisat data
                List<ExcelGlobalCotisatData> globalCotisat = new List<ExcelGlobalCotisatData>();
                if (isGlobalEnt)
                    globalCotisat = CotisatSante.GetCotisatGlobalEntData(years, parentCompanyList);
                else
                    globalCotisat = CotisatSante.GetCotisatGlobalSubsidData(years, subsidiaryList);


                //get all Cadenciers
                List<string> assureurs = globalPresta.Select(x => x.Assureur).Distinct().ToList();
                List<Cadencier> cadencierAll = new List<Cadencier>();
                List<Cadencier> cadencierForAssureur = new List<Cadencier>();
                cadencierAll = Cadencier.GetCadencierForAssureur(C.cDEFAULTASSUREUR);

                foreach (string assurName in assureurs)
                {
                    if (assurName != C.cDEFAULTASSUREUR)
                    {
                        cadencierForAssureur = Cadencier.GetCadencierForAssureur(assurName);
                        cadencierAll.AddRange(cadencierForAssureur);
                    }
                }

                //get coefCad
                Dictionary<int, double> dictYearCoef = new Dictionary<int, double>();                
                foreach (int year in years)
                {
                    DateTime dateDebutSurv = new DateTime(year, 1, 1);
                    DateTime dateFinSurv = new DateTime(year, 12, 31);
                    //### correct the following: assureurs[0]
                    double coefCad = GetCoefCadencier(year, dateArret, dateDebutSurv, dateFinSurv, cadencierAll, assureurs[0]);

                    dictYearCoef.Add(year, coefCad);
                }


                //merge both datasets
                foreach (ExcelGlobalPrestaData dat in globalPresta)
                {
                    ExcelGlobalCotisatData item = null;
                    if (isGlobalEnt)
                        item = globalCotisat.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.YearSurv == dat.YearSurv);
                    else
                        item = globalCotisat.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.Subsid == dat.Subsid && i.YearSurv == dat.YearSurv);

                    dat.CotBrut = 0;
                    dat.CotNet = 0;
                    if (item != null)
                    {
                        dat.CotBrut = item.CotisatBrute.HasValue ? item.CotisatBrute.Value : 0;
                        dat.CotNet = item.Cotisat.HasValue ? item.Cotisat.Value : 0;
                    }

                    double cotNet = dat.CotNet;
                    double cotBrut = dat.CotBrut;

                    //calculate all remaining fields
                    //calculate new dates for dateDebut == 1/1/YearSurv && dateFin == 31/12/YearSurv
                    DateTime dateDebutSurv = new DateTime(dat.YearSurv, 1, 1);
                    DateTime dateFinSurv = new DateTime(dat.YearSurv, 12, 31);
                    double presta = dat.RNous.HasValue ? dat.RNous.Value : 0;
                    double coeffCad = GetCoefCadencier(dat.YearSurv, dateArret, dateDebutSurv, dateFinSurv, cadencierAll, dat.Assureur);

                    double provision = coeffCad * presta;
                    provision = Math.Round(provision, 2);

                    //double cotNet = 0;                    
                    //if ((1 + (TaxDef / 100)) != 0)
                    //    cotNet = (dat.CotBrut * (1 - (TaxAct / 100))) / (1 + (TaxDef / 100));
                    cotNet = Math.Round(cotNet, 2);

                    double ratio = 0;
                    if (cotNet != 0)
                        ratio = (presta + provision) / cotNet;

                    double gainLoss = cotNet - presta - provision;

                    dat.RNous = presta;
                    dat.Coef = coeffCad;
                    dat.Provisions =provision;
                    //dat.CotNet = cotNet;
                    dat.Ratio = ratio;
                    dat.GainLoss = gainLoss;                    
                }

                //Some values from the Cot table may be missing => because we don't have a corresponding entry in the Presta table for certain PK's (Assur-Comp-Year...)
                //we need to add those missing values from the Cot Table
                foreach (ExcelGlobalCotisatData dat in globalCotisat)
                { 
                    ExcelGlobalPrestaData prestaLine = null;
                    if (isGlobalEnt)
                        prestaLine = globalPresta.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.YearSurv == dat.YearSurv);
                    else
                        prestaLine = globalPresta.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.Subsid == dat.Subsid && i.YearSurv == dat.YearSurv);

                    if (prestaLine == null)
                    {
                        double cotNet = 0;
                        double cotBrut = 0;

                        ExcelGlobalPrestaData item = new ExcelGlobalPrestaData();
                        item.Assureur = dat.Assureur;
                        item.Company = dat.Company;
                        item.Subsid = dat.Subsid;
                        item.YearSurv = dat.YearSurv;
                        item.CotBrut = dat.CotisatBrute.HasValue ? dat.CotisatBrute.Value : 0;
                        item.CotNet = dat.Cotisat.HasValue ? dat.Cotisat.Value : 0;

                        cotNet = item.CotNet;
                        cotBrut = item.CotBrut;

                        //if ((1 + (TaxDef / 100)) != 0)
                        //    cotNet = (item.CotBrut * (1 - (TaxAct / 100))) / (1 + (TaxDef / 100));
                        cotNet = Math.Round(cotNet, 2);
                        item.CotNet = cotNet;

                        double gainLoss = cotNet - (item.RNous.HasValue ? item.RNous.Value : 0) - item.Provisions;
                        item.GainLoss = gainLoss;

                        DateTime dateDebutSurv = new DateTime(dat.YearSurv, 1, 1);
                        DateTime dateFinSurv = new DateTime(dat.YearSurv, 12, 31);
                        double coeffCad = GetCoefCadencier(dat.YearSurv, dateArret, dateDebutSurv, dateFinSurv, cadencierAll, dat.Assureur);
                        item.Coef = coeffCad;

                        globalPresta.Add(item);
                    }
                }

                //CUMUL
                List<ExcelGlobalPrestaData> globalCotisatCumul = globalPresta
                   .GroupBy(p => new { p.Assureur, p.YearSurv })
                   .Select(g => new ExcelGlobalPrestaData
                   {
                       Assureur = g.Key.Assureur,
                       Company = "",
                       Subsid = "",
                       YearSurv = g.Key.YearSurv,
                       RNous = g.Sum(i => i.RNous),
                       Provisions = g.Sum(i => i.Provisions),
                       CotBrut = g.Sum(i => i.CotBrut),
                       TaxActive = TaxAct.ToString(),
                       TaxDefault = TaxDef.ToString(),
                       TaxTotal = "",
                       CotNet = g.Sum(i => i.CotNet),
                       Ratio = ((g.Sum(i => i.RNous) + g.Sum(i => i.Provisions) ) / g.Sum(i => i.CotNet)).Value,
                       GainLoss = g.Sum(i => i.GainLoss),
                       Coef = dictYearCoef[g.Key.YearSurv],
                       FR = g.Sum(i => i.FR),
                       RSS = g.Sum(i => i.RSS),
                       RAnnexe = g.Sum(i => i.RAnnexe),
                       DateArret = dateArret
                   })
                   .OrderBy(ga => ga.YearSurv)
                   .ToList();

                //double? total = globalPresta.Sum(p => p.FR);

                //create DATA_CUMUL table
                foreach (ExcelGlobalPrestaData prest in globalCotisatCumul)
                {
                    //calculate CotNet
                    //double cotNet = 0;
                    //if ((1 + (TaxDef / 100)) != 0)
                    //    cotNet = (prest?.CotBrut ?? 0) * (1 - (TaxDef / 100)) / (1 + (TaxAct / 100));

                    double ratio = 0;
                    if (prest.CotNet != 0)
                        ratio = (prest.Provisions + (prest.RNous.HasValue ? prest.RNous.Value : 0 )) / prest.CotNet;

                    double? tauxChargement = 0;
                    if(prest?.CotBrut != 0)
                    {
                        tauxChargement = 1 - (prest?.CotNet / prest?.CotBrut);
                    }

                    //double gainLoss = cotNet - (prest.RNous.HasValue ? prest.RNous.Value : 0) - prest.Provisions;

                    DataRow newRow = globalTableCumul.NewRow();

                    newRow["Assureur"] = prest.Assureur;
                    newRow["Company"] = "";
                    newRow["Subsid"] = "";

                    newRow["YearSurv"] = prest.YearSurv;
                    newRow["FR"] = prest?.FR ?? 0;
                    newRow["RSS"] = prest?.RSS ?? 0;
                    newRow["RAnnexe"] = prest.RAnnexe ?? 0;
                    newRow["RNous"] = prest?.RNous ?? 0;
                    newRow["Provisions"] = Math.Round(prest.Provisions, 2);
                    newRow["CoeffCad"] = prest.Coef;
                    newRow["CotBrut"] = prest?.CotBrut ?? 0;
                    //newRow["TaxTotal"] = string.Format("{0:0.00} %", (TaxDef + TaxAct));
                    //newRow["TaxTotal"] = string.Format("{0:0.0000} %", tauxChargement);
                    //newRow["TaxDefault"] = string.Format("{0:0.00} %", TaxDef);
                    //newRow["TaxActive"] = string.Format("{0:0.00} %", TaxAct);
                    //newRow["CotNet"] = Math.Round(cotNet, 2);
                    newRow["CotNet"] = prest.CotNet;
                    newRow["Ratio"] = Math.Round(ratio, 4);                    
                    newRow["GainLoss"] = prest.GainLoss;
                    newRow["DateArret"] = dateArret;

                    globalTableCumul.Rows.Add(newRow);
                }
                

                //create DATA table
                foreach (ExcelGlobalPrestaData prest in globalPresta)
                {
                    double? tauxChargement = 0;
                    if (prest?.CotBrut != 0)
                    {
                        tauxChargement = 1 - (prest?.CotNet / prest?.CotBrut);
                    }

                    DataRow newRow = globalTable.NewRow();

                    newRow["Assureur"] = prest.Assureur;
                    newRow["Company"] = prest.Company;

                    if (!isGlobalEnt)
                        newRow["Subsid"] = prest.Subsid;
                    else
                        newRow["Subsid"] = "";

                    newRow["YearSurv"] = prest.YearSurv;
                    newRow["FR"] = prest?.FR ?? 0;
                    newRow["RSS"] = prest?.RSS ?? 0; 
                    newRow["RAnnexe"] = prest.RAnnexe ?? 0;
                    newRow["RNous"] = prest?.RNous ?? 0;
                    newRow["Provisions"] = Math.Round(prest.Provisions, 2);
                    newRow["CoeffCad"] = prest?.Coef ?? 0;
                    newRow["CotBrut"] = prest?.CotBrut ?? 0;
                    //newRow["TaxTotal"] = string.Format("{0:0.00} %", (TaxDef + TaxAct));
                    //newRow["TaxTotal"] = string.Format("{0:0.0000} %", tauxChargement);
                    //newRow["TaxDefault"] = string.Format("{0:0.00} %", TaxDef);
                    //newRow["TaxActive"] = string.Format("{0:0.00} %", TaxAct);
                    newRow["CotNet"] = Math.Round(prest.CotNet, 2);
                    newRow["Ratio"] = Math.Round(prest.Ratio, 4);
                    newRow["GainLoss"] = prest.GainLoss;
                    newRow["DateArret"] = dateArret;
                 
                    globalTable.Rows.Add(newRow);
                }

                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELGLOBAL].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);
                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELGLOBAL];
                    ws.Cells["A2"].LoadFromDataTable(globalTable, false);

                    pck.Workbook.Worksheets[C.cEXCELGLOBALCUMUL].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);
                    ws = pck.Workbook.Worksheets[C.cEXCELGLOBALCUMUL];
                    ws.Cells["A2"].LoadFromDataTable(globalTableCumul, false);

                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillGlobalEntSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillGlobalSheetPrev(FileInfo excelFilePath, string companyList, string subsidList, DateTime debutPeriod,
           DateTime finPeriod, DateTime dateArret, C.eReportTypes reportType, double TaxDef, double TaxAct, double TaxPer)
        {
            try
            {
                bool isGlobalEnt = reportType == C.eReportTypes.GlobalEnt ? true : false;
                List<string> parentCompanyList = Regex.Split(companyList, C.cVALSEP).ToList();
                List<string> subsidiaryList = Regex.Split(subsidList, C.cVALSEP).ToList();
                DataTable globalTable = CreateGlobalTablePrev(reportType);
                DataTable globalTableCumul = CreateGlobalTablePrev(reportType);

                //get years && coefCad                
                List<int> years = new List<int>();
                for (int i = 0; i <= finPeriod.Year - debutPeriod.Year; i++)
                {
                    years.Add(debutPeriod.Year + i);
                }

                //get presta data
                List<ExcelGlobalDecompteData> globalDecompte = new List<ExcelGlobalDecompteData>();
                if (isGlobalEnt)
                    globalDecompte = DecomptePrev.GetDecompteGlobalEntData(years, parentCompanyList);
                else
                    globalDecompte = DecomptePrev.GetDecompteGlobalSubsidData(years, subsidiaryList);

                //get cotisat data
                List<ExcelGlobalCotisatData> globalCotisat = new List<ExcelGlobalCotisatData>();
                if (isGlobalEnt)
                    globalCotisat = CotisatPrev.GetCotisatGlobalEntDataPrev(years, parentCompanyList);
                else
                    globalCotisat = CotisatPrev.GetCotisatGlobalSubsidDataPrev(years, subsidiaryList);
                
                //merge both datasets
                foreach (ExcelGlobalDecompteData dat in globalDecompte)
                {
                    ExcelGlobalCotisatData item = null;
                    if (isGlobalEnt)
                        item = globalCotisat.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.YearSurv == dat.YearSurv);
                    else
                        item = globalCotisat.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.Subsid == dat.Subsid && i.YearSurv == dat.YearSurv);

                    if (item != null)
                    {
                        dat.CotBrute = item.CotisatBrute.HasValue ? item.CotisatBrute.Value : 0;
                        dat.CotNet = item.Cotisat.HasValue ? item.Cotisat.Value : 0;
                    }

                    //calculate all remaining fields
                    //calculate new dates for dateDebut == 1/1/YearSurv && dateFin == 31/12/YearSurv
                    DateTime dateDebutSurv = new DateTime(dat.YearSurv, 1, 1);
                    DateTime dateFinSurv = new DateTime(dat.YearSurv, 12, 31);
                    double presta = dat.RNous.HasValue ? dat.RNous.Value : 0;
                    double coeffCad = 0;

                    double provision = 0;
                    provision = Math.Round(provision, 2);

                    double cotNet = 0;
                    cotNet = dat.CotNet;

                    double cotBrute = 0;
                    cotBrute = dat.CotBrute;
                    //if ((1 + (TaxDef / 100)) != 0)
                    //    cotNet = (dat.CotBrut * (1 - (TaxAct / 100))) / (1 + (TaxDef / 100));
                    cotNet = Math.Round(cotNet, 2);
                    cotBrute = Math.Round(cotBrute, 2);

                    double ratio = 0;
                    if (cotNet != 0)
                        ratio = (presta + provision) / cotNet;

                    double gainLoss = cotNet - presta - provision;

                    dat.RNous = presta;
                    dat.Coef = 0;
                    dat.Provisions = provision;
                    dat.CotNet = cotNet;
                    dat.CotBrute = cotBrute;
                    dat.Ratio = ratio;
                    dat.GainLoss = gainLoss;
                }

                //Some values from the Cot table may be missing => because we don't have a corresponding entry in the Presta table for certain PK's (Assur-Comp-Year...)
                //we need to add those missing values from the Cot Table
                foreach (ExcelGlobalCotisatData dat in globalCotisat)
                {
                    ExcelGlobalDecompteData decompteLine = null;
                    if (isGlobalEnt)
                        decompteLine = globalDecompte.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.YearSurv == dat.YearSurv);
                    else
                        decompteLine = globalDecompte.FirstOrDefault(i => i.Assureur == dat.Assureur && i.Company == dat.Company && i.Subsid == dat.Subsid && i.YearSurv == dat.YearSurv);

                    if (decompteLine == null)
                    {
                        ExcelGlobalDecompteData item = new ExcelGlobalDecompteData();
                        item.Assureur = dat.Assureur;
                        item.Company = dat.Company;
                        item.Subsid = dat.Subsid;
                        item.YearSurv = dat.YearSurv;
                        item.CotNet = dat.Cotisat.HasValue ? dat.Cotisat.Value : 0;
                        item.CotBrute = dat.CotisatBrute.HasValue ? dat.CotisatBrute.Value : 0;

                        double cotNet = 0;
                        cotNet = item.CotNet;
                        double cotBrute = 0;
                        cotBrute = item.CotBrute;
                        //if ((1 + (TaxDef / 100)) != 0)
                        //    cotNet = (item.CotBrut * (1 - (TaxAct / 100))) / (1 + (TaxDef / 100));
                        cotNet = Math.Round(cotNet, 2);
                        item.CotNet = cotNet;
                        cotBrute = Math.Round(cotBrute, 2);
                        item.CotBrute = cotBrute;

                        double gainLoss = cotNet - (item.RNous.HasValue ? item.RNous.Value : 0) - item.Provisions;
                        item.GainLoss = gainLoss;                        

                        DateTime dateDebutSurv = new DateTime(dat.YearSurv, 1, 1);
                        DateTime dateFinSurv = new DateTime(dat.YearSurv, 12, 31);
                        double coeffCad = 0;
                        item.Coef = coeffCad;

                        globalDecompte.Add(item);
                    }
                }

                //CUMUL
                List<ExcelGlobalDecompteData> globalCotisatCumul = globalDecompte
                   .GroupBy(p => new { p.Assureur, p.YearSurv })
                   .Select(g => new ExcelGlobalDecompteData
                   {
                       Assureur = g.Key.Assureur,
                       Company = "",
                       Subsid = "",
                       YearSurv = g.Key.YearSurv,
                       RNous = g.Sum(i => i.RNous),
                       Provisions = g.Sum(i => i.Provisions),
                       CotBrute = g.Sum(i => i.CotBrute),
                       TaxActive = TaxAct.ToString(),
                       TaxDefault = TaxDef.ToString(),
                       TaxTotal = "",
                       CotNet = g.Sum(i => i.CotNet),
                       Ratio = ((g.Sum(i => i.RNous) + g.Sum(i => i.Provisions)) / g.Sum(i => i.CotNet)).Value,
                       GainLoss = g.Sum(i => i.GainLoss),
                       Coef = 0,
                       FR = 0,
                       RSS = 0,
                       RAnnexe = 0,
                       DateArret = dateArret
                   })
                   .OrderBy(ga => ga.YearSurv)
                   .ToList();

                //double? total = globalPresta.Sum(p => p.FR);

                //create DATA_CUMUL table
                //Assureur, Company, Subsid, YearSurv, TypeSinistre, Prestations, Provisions, CotBrut, TauxChargement, CotNet, Ratio, GainLoss
                foreach (ExcelGlobalDecompteData decompte in globalCotisatCumul)
                {
                    //calculate CotNet
                    //double cotNet = 0;
                    //if ((1 + (TaxDef / 100)) != 0)
                    //    cotNet = (prest?.CotBrut ?? 0) * (1 - (TaxDef / 100)) / (1 + (TaxAct / 100));

                    double ratio = 0;
                    if (decompte.CotNet != 0)
                        ratio = (decompte.Provisions + (decompte.RNous.HasValue ? decompte.RNous.Value : 0)) / decompte.CotNet;

                    double tauxChargement = 0;
                    if (decompte.CotBrute != 0)
                    {
                        tauxChargement = 1 - (decompte.CotNet / decompte.CotBrute);
                    }

                    //double gainLoss = cotNet - (prest.RNous.HasValue ? prest.RNous.Value : 0) - prest.Provisions;

                    DataRow newRow = globalTableCumul.NewRow();

                    newRow["Assureur"] = decompte.Assureur;
                    newRow["Company"] = "";
                    newRow["Subsid"] = "";
                    newRow["YearSurv"] = decompte.YearSurv;
                    newRow["TypeSinistre"] = "";
                    newRow["Prestations"] = decompte?.RNous ?? 0;
                    newRow["Provisions"] = Math.Round(decompte.Provisions, 2);
                    newRow["CotBrute"] = decompte?.CotBrute ?? 0;
                    newRow["TauxChargement"] = string.Format("{0:0.0000} %", tauxChargement);
                    //newRow["CotNet"] = Math.Round(cotNet, 2);
                    newRow["CotNet"] = decompte?.CotNet ?? 0;
                    newRow["Ratio"] = Math.Round(ratio, 4);
                    newRow["GainLoss"] = decompte.GainLoss;
                    newRow["DateArret"] = dateArret;

                    globalTableCumul.Rows.Add(newRow);
                }

                //create DATA table
                foreach (ExcelGlobalDecompteData decompte in globalDecompte)
                {
                    double tauxChargement = 0;
                    if (decompte.CotBrute != 0)
                    {
                        tauxChargement = 1 - (decompte.CotNet / decompte.CotBrute);
                    }

                    DataRow newRow = globalTable.NewRow();

                    newRow["Assureur"] = decompte.Assureur;
                    newRow["Company"] = decompte.Company;

                    if (!isGlobalEnt)
                        newRow["Subsid"] = decompte.Subsid;
                    else
                        newRow["Subsid"] = "";

                    newRow["YearSurv"] = decompte.YearSurv;
                    newRow["TypeSinistre"] = "";
                    newRow["Prestations"] = decompte?.RNous ?? 0;
                    newRow["Provisions"] = Math.Round(decompte.Provisions, 2);
                    newRow["CotBrute"] = decompte?.CotBrute ?? 0;
                    newRow["TauxChargement"] = string.Format("{0:0.0000} %", tauxChargement);
                    newRow["CotNet"] = decompte?.CotNet ?? 0;
                    newRow["Ratio"] = Math.Round(decompte.Ratio, 4);
                    newRow["GainLoss"] = decompte.GainLoss;
                    newRow["DateArret"] = dateArret;

                    globalTable.Rows.Add(newRow);
                }

                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELGLOBAL].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);
                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELGLOBAL];
                    ws.Cells["A2"].LoadFromDataTable(globalTable, false);

                    pck.Workbook.Worksheets[C.cEXCELGLOBALCUMUL].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);
                    ws = pck.Workbook.Worksheets[C.cEXCELGLOBALCUMUL];
                    ws.Cells["A2"].LoadFromDataTable(globalTableCumul, false);

                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillGlobalEntSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillPrevSheet(FileInfo excelFilePath, string assurNameList, string parentCompanyNameList, string companyNameList, string contrNameList, string college, 
            DateTime debutPeriod, DateTime finPeriod, DateTime dateArret, int yearsToCalc)
        {
            try
            {
                List<AgeLabel> ageLabels = AgeLabel.GetAgeLabels();
                List<TypePrevoyance> typePref = TypePrevoyance.GetTypePrev();

                List<string> contrList = Regex.Split(contrNameList, C.cVALSEP).ToList();
                List<string> parentCompanyList = Regex.Split(parentCompanyNameList, C.cVALSEP).ToList();
                List<string> companyList = Regex.Split(companyNameList, C.cVALSEP).ToList();
                List<string> assurList = Regex.Split(assurNameList, C.cVALSEP).ToList();

                List<SinistrePrev> mySinistreData = new List<SinistrePrev>();
                List<SinistrePrev> yearSinistreData = new List<SinistrePrev>();

                //certain report templates will require data for more than 1 year, take this into account  
                DateTime debutNew;
                DateTime finNew;

                //recherche sur la période saisie sans prendre en compte le recul (base de données: yearsToCalc) : RS 12/11/2020
                int years = 0;
                //for (int years = 0; years < yearsToCalc; years++)
                //{
                    debutNew = new DateTime(debutPeriod.Year - years, debutPeriod.Month, debutPeriod.Day);
                    finNew = new DateTime(finPeriod.Year - years, finPeriod.Month, finPeriod.Day);

                    yearSinistreData = SinistrePrev.GetSinistresForContracts(assurList, parentCompanyList, companyList, contrList, college, debutNew, finNew, dateArret);
                    mySinistreData.AddRange(yearSinistreData);
                //}


                //transform the data
                DataTable prevTable = new DataTable();

                DataColumn company = new DataColumn("CLIENT", typeof(string));
                DataColumn contract = new DataColumn("CONTRAT", typeof(string));
                DataColumn codcol = new DataColumn("LIBCOL", typeof(string));
                DataColumn dateSinistre = new DataColumn("Date Sinistre", typeof(DateTime));
                DataColumn yearSinistre = new DataColumn("ANNESIN", typeof(int));
                DataColumn typePrev = new DataColumn("TYPE PREV", typeof(string));
                DataColumn birthdate = new DataColumn("Date de Naissance", typeof(DateTime));
                DataColumn age = new DataColumn("AGE", typeof(int));
                DataColumn ageRange = new DataColumn("TRANCHE AGE", typeof(string));
                DataColumn sex1 = new DataColumn("SEXE", typeof(string));
                DataColumn dateClosure = new DataColumn("DATECLOTURE", typeof(DateTime));
                DataColumn motifClosure = new DataColumn("MOTIFCLO", typeof(string));
                DataColumn duration = new DataColumn("DUREE", typeof(int));
                DataColumn ageEntreeInval = new DataColumn("AGE ENTREE INVAL", typeof(int));
                DataColumn nbJoursInval = new DataColumn("Nbj Inval", typeof(int));
                DataColumn prestaTotal = new DataColumn("PRESTATION", typeof(double));
                DataColumn provision = new DataColumn("PROVISION", typeof(decimal));   
                DataColumn dossier1 = new DataColumn("DOSSIER", typeof(string));
                DataColumn lastname = new DataColumn("Nom", typeof(string));
                DataColumn firstname = new DataColumn("Prénom", typeof(string));
                DataColumn sexHF = new DataColumn("Sexe (H/F)", typeof(string));
                DataColumn dateResiliat = new DataColumn("Date de Resiliation", typeof(DateTime));
                DataColumn birthdate2 = new DataColumn("Date de Naissance2", typeof(DateTime));
                DataColumn dateIncap = new DataColumn("Date d'entrée en Incapacité", typeof(DateTime));
                DataColumn franchiseIncap = new DataColumn("Franchise en Mois pour l'incapacité", typeof(int)); //???
                DataColumn durationIncap = new DataColumn("Durée de la prestation d'incapacité en mois", typeof(int));
                DataColumn incapYesNo = new DataColumn("incapacité avec Passage en invalidité Oui/Non (O/N)", typeof(string));
                DataColumn dateInval = new DataColumn("Date d'entrée en Invalidité", typeof(DateTime));
                DataColumn prestaIncapInval = new DataColumn("Prestation annualisée pour l'incapacité ou l'invalidité", typeof(decimal));
                DataColumn franchiseInval = new DataColumn("Franchise en années pour l'invalidité", typeof(int));
                DataColumn capitalDeces = new DataColumn("Montant du Capital Décès", typeof(decimal));
                DataColumn categoryInval = new DataColumn("Catégorie d'invalidité(1 , 2 ou 3)", typeof(int));
                DataColumn dateMax = new DataColumn("Date Max", typeof(DateTime));


                prevTable.Columns.AddRange(new DataColumn[] { company, contract, codcol, dateSinistre, yearSinistre, typePrev, birthdate, age, ageRange, sex1,
                    dateClosure, motifClosure, duration, ageEntreeInval, nbJoursInval, prestaTotal, provision, dossier1, lastname, firstname, sexHF, dateResiliat,
                    birthdate2, dateIncap, franchiseIncap, durationIncap, incapYesNo, dateInval, prestaIncapInval, franchiseInval, capitalDeces, categoryInval, dateMax});


                //get a list of all dossiers from mySinistreData
                List<string> dossierList = mySinistreData.Select(s => s.Dossier).ToList();

                //get all decomptes for all dossiers listed in dossierList
                List<DecomptePrevReduced> decPrevReduced = DecomptePrev.GetDecomptesForDossierList(dossierList, dateArret);   


                foreach (SinistrePrev sin in mySinistreData)
                {
                    double presta = 0;
                    DataRow newRow = prevTable.NewRow();
                    string dossier = sin.Dossier;

                    bool motifClotureIsInval = false;

                    if(sin.MotifClo !=null)
                        motifClotureIsInval = sin.MotifClo.ToLower().Contains("inval") ? true : false;

                    newRow["CLIENT"] = sin.Company;
                    newRow["CONTRAT"] = sin.ContractId;                    
                    newRow["LIBCOL"] = sin.CodeCol;                    

                    newRow["Date Sinistre"] = sin.DateSinistre.HasValue ? sin.DateSinistre.Value : (object)DBNull.Value;
                    newRow["ANNESIN"] = sin.DateSinistre.HasValue ? sin.DateSinistre.Value.Year : (object)DBNull.Value;

                    //get type of prev from ref table
                    // AM le 14 01 2019
                    //TypePrevoyance tp = typePref.Find(p => p.CodeSinistre == sin.CauseSinistre);
                    TypePrevoyance tp = typePref.Find(p => p.CodeSinistre == sin.NatureSinistre);
                    
                    if (tp != null)
                        newRow["TYPE PREV"] = tp.LabelSinistre;

                    newRow["Date de Naissance"] = sin.Birthdate.HasValue ? sin.Birthdate.Value : (object)DBNull.Value;

                    int personAge = 0;
                    if (sin.Birthdate.HasValue && sin.DateSinistre.HasValue)
                    {
                        personAge = sin.DateSinistre.Value.Year - sin.Birthdate.Value.Year;
                        //if (sin.DateSinistre.Value < sin.Birthdate.Value.AddYears(personAge)) personAge--;
                    }
                    newRow["AGE"] = personAge;

                    //get age label from ref table
                    AgeLabel al = ageLabels.Find(a => a.Age == personAge);
                    if (al != null)
                        newRow["TRANCHE AGE"] = al.Label;

                    string title = sin.Title != null ? sin.Title : "mr";
                    if (title.ToLower() == "mr" || title.ToLower()=="m") title = "Masculin"; else title = "Féminin";

                    newRow["SEXE"] = title;

                    newRow["DATECLOTURE"] = sin.DateClo.HasValue ? sin.DateClo.Value : (object)DBNull.Value;
                    newRow["MOTIFCLO"] = sin.MotifClo;

                    //Nbj Incap
                    int days = 0;
                    if (sin.DateClo.HasValue && sin.DateSinistre.HasValue)
                    {
                        days = (int)(sin.DateClo.Value - sin.DateSinistre.Value).TotalDays;                        
                    }
                    // AM le 14 01 2019
                    if (days < 0)
                    {
                        days = 0;
                    }

                    newRow["DUREE"] = days;

                    //AGE ENTREE INVAL
                    personAge = 0;
                    // AM le 14 01 2019
                    //if (sin.DateClo.HasValue && sin.Birthdate.HasValue)
                    //{
                    //    personAge = sin.DateClo.Value.Year - sin.Birthdate.Value.Year;
                    //    //if (sin.DateClo.Value < sin.Birthdate.Value.AddYears(personAge)) personAge--;
                    //}
                    // AM le 14 01 2019
                    if (sin.DateSinistre.HasValue && sin.Birthdate.HasValue)
                    {
                        personAge = sin.DateSinistre.Value.Year - sin.Birthdate.Value.Year;
                        //if (sin.DateClo.Value < sin.Birthdate.Value.AddYears(personAge)) personAge--;
                    }

                    if (motifClotureIsInval)
                        newRow["AGE ENTREE INVAL"] = personAge;
                    else
                        newRow["AGE ENTREE INVAL"] = (object)DBNull.Value;

                    //Nbj Inval
                    days = 0;
                    // AM le 14 01 2019 ok
                    //if (sin.DateClo.HasValue )
                    //{
                    //    days = (int)(dateArret - sin.DateClo.Value).TotalDays;
                    //}
                    //// AM le 14 01 2019
                    if (sin.DateClo.HasValue && sin.DateSinistre.HasValue)
                    {
                        days = (int)(sin.DateClo.Value - sin.DateSinistre.Value).TotalDays;
                    }
                    if (days < 0)
                    {
                        days = 0;
                    }
                    if (motifClotureIsInval)
                        newRow["Nbj Inval"] = days;
                    else
                        newRow["Nbj Inval"] = (object)DBNull.Value;


                    //calculate from Decompte Table                   

                    //decimal presta1 = DecomptePrev.GetSumPrestaForDossier(dossier, dateArret);
                    //presta = DecomptePrev.GetSumPrestaForDossierFromSP(dossier, dateArret);
                    var res = decPrevReduced.Where(d => d.Dossier == dossier && d.DatePayement < dateArret).Sum(d => d.Total);                    
                    presta = res.HasValue ? res.Value : 0;

                    newRow["PRESTATION"] = presta; 
                    
                    newRow["PROVISION"] = (object)DBNull.Value;

                    newRow["DOSSIER"] = dossier;

                    newRow["Nom"] = sin.Lastname;
                    newRow["Prénom"] = sin.Firstname;

                    if (title.ToLower() == "masculin") title = "H"; else title = "F";
                    newRow["Sexe (H/F)"] = title;                    
                    
                    newRow["Date de Resiliation"] = dateArret;                     
                    newRow["Date de Naissance2"] = sin.Birthdate.HasValue ? sin.Birthdate.Value : (object)DBNull.Value;
                    newRow["Date d'entrée en Incapacité"] = sin.DateSinistre.HasValue ? sin.DateSinistre.Value : (object)DBNull.Value;                                        
                    newRow["Franchise en Mois pour l'incapacité"] = 0;
                    newRow["Durée de la prestation d'incapacité en mois"] = 36;
                    newRow["incapacité avec Passage en invalidité Oui/Non (O/N)"] = "O";

                    
                    if (motifClotureIsInval)
                        // AM le 14 01 2019
                        //newRow["Date d'entrée en Invalidité"] = sin.DateClo.HasValue ? sin.DateClo.Value : (object)DBNull.Value;
                        // AM le 14 01 2019
                        newRow["Date d'entrée en Invalidité"] = sin.DateSinistre.HasValue ? sin.DateSinistre.Value : (object)DBNull.Value; 
                    else
                        newRow["Date d'entrée en Invalidité"] = (object)DBNull.Value;

                    //presta = DecomptePrev.GetSumPrestaAnnualized(dossier, dateArret);                    
                    //presta = DecomptePrev.GetSumPrestaAnnualizedFromSP(dossier, dateArret);                    
                    //res = decPrevReduced.Where(d => d.Dossier == dossier && d.DatePayement < dateArret && d.FinSin.HasValue && d.DebSin.HasValue && d.Total.HasValue)
                    //    .Sum(d => d.Total / (int)((d.FinSin.Value - d.DebSin.Value).TotalDays + 1) * 365);
                    //presta = res.HasValue ? res.Value : 0;

                    double? sumTotals = decPrevReduced.Where(d => d.Dossier == dossier && d.DatePayement < dateArret && d.FinSin.HasValue && d.DebSin.HasValue && d.Total.HasValue)
                        .Sum(d => d.Total );

                    int sumDays = decPrevReduced.Where(d => d.Dossier == dossier && d.DatePayement < dateArret && d.FinSin.HasValue && d.DebSin.HasValue && d.Total.HasValue)
                        .Sum(d => (int)((d.FinSin.Value - d.DebSin.Value).TotalDays + 1));

                    if (sumTotals.HasValue && sumDays > 0 && sumTotals > 0)
                        presta = sumTotals.Value / sumDays * 365;                    

                    newRow["Prestation annualisée pour l'incapacité ou l'invalidité"] = presta;
                   
                    
                    newRow["Franchise en années pour l'invalidité"] = 0;                    
                    newRow["Montant du Capital Décès"] = 0;                    
                    newRow["Catégorie d'invalidité(1 , 2 ou 3)"] = 1;
                    
                    //var maxDate = DecomptePrev.GetDateMaxForDossier(dossier, dateArret);                    
                    //var maxDate = DecomptePrev.GetDateMaxForDossierFromSP(dossier, dateArret);

                    var maxDate = decPrevReduced.Where(d => d.Dossier == dossier && d.DatePayement < dateArret).Max(d => d.FinSin);
                    if (maxDate.HasValue)
                        newRow["Date Max"] = maxDate;
                    else
                        newRow["Date Max"] = (object)DBNull.Value;

                    prevTable.Rows.Add(newRow);
                } 


                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    try
                    {                        
                        pck.Workbook.Worksheets[C.cEXCELPREV].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);
                    }
                    catch (Exception ex2) { }

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELPREV];
                    ws.Cells["A2"].LoadFromDataTable(prevTable, false);
                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillPrevSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillDemoSheet(FileInfo excelFilePath, string assurNameList, string parentCompanyNameList, string companyNameList, string contrNameList, 
            DateTime debutPeriod, DateTime finPeriod, DateTime dateArret, int yearsToCalc, bool reportWithOption)
        {
            try
            {
                List<CDemoData> myDemoData = new List<CDemoData>();
                List<CDemoData> yearDemoData;

                List<CDemoDataWithoutOption> myDemoDataWO = new List<CDemoDataWithoutOption>();
                List<CDemoDataWithoutOption> yearDemoDataWO;

                //certain report templates will require data for more than 1 year, take this into account
                DateTime debutNew;
                DateTime finNew;

                int years = 0;
                //for (int years = 0; years < yearsToCalc; years++)
                //{
                    debutNew = new DateTime(debutPeriod.Year - years, debutPeriod.Month, debutPeriod.Day);
                    finNew = new DateTime(finPeriod.Year - years, finPeriod.Month, finPeriod.Day);

                    if (!reportWithOption)
                    {
                        //standard SP without option data
                        yearDemoDataWO = Demography.GetDemoDataWithoutOptionFromSP(assurNameList, parentCompanyNameList, companyNameList, contrNameList, debutNew, finNew);
                        myDemoDataWO.AddRange(yearDemoDataWO);
                    }
                    else
                    {
                        yearDemoData = Demography.GetDemoDataFromSP(assurNameList, parentCompanyNameList, companyNameList, contrNameList, debutNew, finNew);
                        myDemoData.AddRange(yearDemoData);
                    }
                //}                               

                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELDEMO].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELDEMO];

                    if (!reportWithOption)
                    {
                        ws.Cells["A2"].LoadFromCollection(myDemoDataWO);
                    }
                    else
                    {
                        ws.Cells["A2"].LoadFromCollection(myDemoData);
                    }

                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillDemoSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillCotSheet(FileInfo excelFilePath, string assurNameList, string parentCompanyNameList, string companyNameList, string contrNameList, string college, DateTime debutPeriod, 
            DateTime finPeriod, DateTime dateArret, int yearsToCalc, C.eReportTemplateTypes templateType)
        {
            try
            {                
                List<string> contrList = Regex.Split(contrNameList, C.cVALSEP).ToList();
                List<string> parentCompanyList = Regex.Split(parentCompanyNameList, C.cVALSEP).ToList();
                List<string> companyList = Regex.Split(companyNameList, C.cVALSEP).ToList();
                List<string> assurList = Regex.Split(assurNameList, C.cVALSEP).ToList();

                List<CotisatSante> myCotDataSante = new List<CotisatSante>();
                List<CotisatSante> yearCotDataSante = new List<CotisatSante>();
                List<CotisatPrev> myCotDataPrev = new List<CotisatPrev>();
                List<CotisatPrev> yearCotDataPrev = new List<CotisatPrev>();

                //certain report templates will require data for more than 1 year, take this into account  
                DateTime debutNew;
                DateTime finNew;

                int years = 0;
                //for (int years = 0; years < yearsToCalc; years++)
                //{
                    debutNew = new DateTime(debutPeriod.Year - years, debutPeriod.Month, debutPeriod.Day);
                    finNew = new DateTime(finPeriod.Year - years, finPeriod.Month, finPeriod.Day);

                    if (templateType == C.eReportTemplateTypes.SANTE || templateType == C.eReportTemplateTypes.SANTE_GLOBAL)
                    {
                        yearCotDataSante = CotisatSante.GetCotisationsForContracts(assurList, parentCompanyList, companyList, contrList, college, debutNew, finNew);
                        myCotDataSante.AddRange(yearCotDataSante);
                    }
                    else if ((templateType == C.eReportTemplateTypes.PREV && years == 0) || (templateType == C.eReportTemplateTypes.PREV_GLOBAL && years == 0))
                    {                        
                        yearCotDataPrev = CotisatPrev.GetCotisationsForContracts(assurList, parentCompanyList, companyList, contrList, college, debutNew, finNew);
                        myCotDataPrev.AddRange(yearCotDataPrev);
                    }                    
                //}
                
                //transform the data
                DataTable cotisatTable = new DataTable();

                DataColumn debprime = new DataColumn("DEBPRIME", typeof(DateTime));
                DataColumn finprime = new DataColumn("FINPRIME", typeof(DateTime));
                DataColumn contract = new DataColumn("CONTRAT", typeof(string));
                DataColumn codcol = new DataColumn("CODCOL", typeof(string));
                DataColumn annee = new DataColumn("ANNEE", typeof(int));
                DataColumn cotisat = new DataColumn("COTISATION", typeof(decimal));
                DataColumn famille_college = new DataColumn("FAMILLE_COLLEGE", typeof(string));
                DataColumn cotisatBrute = new DataColumn("COTISATION_BRUTE", typeof(decimal));                

                if (templateType == C.eReportTemplateTypes.SANTE || templateType == C.eReportTemplateTypes.SANTE_GLOBAL)
                {
                    cotisatTable.Columns.AddRange(new DataColumn[] { debprime, finprime, contract, codcol, annee, cotisat, famille_college, cotisatBrute });

                    foreach (CotisatSante cot in myCotDataSante)
                    {
                        DataRow newRow = cotisatTable.NewRow();

                        newRow["DEBPRIME"] = cot.DebPrime.HasValue ? cot.DebPrime.Value.ToShortDateString() : (object)DBNull.Value;
                        newRow["FINPRIME"] = cot.FinPrime.HasValue ? cot.FinPrime.Value.ToShortDateString() : (object)DBNull.Value;
                        newRow["CONTRAT"] = cot.ContractId;
                        newRow["CODCOL"] = cot.CodeCol;
                        newRow["ANNEE"] = cot.Year.HasValue ? cot.Year.Value : 0;
                        newRow["COTISATION"] = cot.Cotisation.HasValue ? cot.Cotisation.Value : 0;
                        newRow["FAMILLE_COLLEGE"] = "";
                        newRow["COTISATION_BRUTE"] = cot.CotisationBrute.HasValue ? cot.CotisationBrute.Value : 0;

                        cotisatTable.Rows.Add(newRow);
                    }
                }
                else if (templateType == C.eReportTemplateTypes.PREV || templateType == C.eReportTemplateTypes.PREV_GLOBAL)
                {
                    cotisatTable.Columns.AddRange(new DataColumn[] { debprime, finprime, contract, codcol, annee, cotisat, cotisatBrute });

                    foreach (CotisatPrev cot in myCotDataPrev)
                    {
                        DataRow newRow = cotisatTable.NewRow();

                        newRow["DEBPRIME"] = cot.DebPrime.HasValue ? cot.DebPrime.Value.ToShortDateString() : (object)DBNull.Value;
                        newRow["FINPRIME"] = cot.FinPrime.HasValue ? cot.FinPrime.Value.ToShortDateString() : (object)DBNull.Value;
                        newRow["CONTRAT"] = cot.ContractId;
                        newRow["CODCOL"] = cot.CodeCol;
                        newRow["ANNEE"] = cot.Year.HasValue ? cot.Year.Value : 0;
                        newRow["COTISATION"] = cot.Cotisation.HasValue ? cot.Cotisation.Value : 0;                        
                        newRow["COTISATION_BRUTE"] = cot.CotisationBrute.HasValue ? cot.CotisationBrute.Value : 0;

                        cotisatTable.Rows.Add(newRow);
                    }
                }

                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELCOT].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELCOT];
                    ws.Cells["A2"].LoadFromDataTable(cotisatTable, false);
                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillCotSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillPrestSheet(FileInfo excelFilePath, CRPlanning crp, List<PrestSante> myPrestData, bool reportWithOption)
        {
            try
            {
                if(reportWithOption)
                    CollectPrestaData2(excelFilePath, crp, myPrestData, C.eExcelSheetPrestaData.Prestation);
                else
                    CollectPrestaData(excelFilePath, crp, myPrestData, C.eExcelSheetPrestaData.Prestation);
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillPrestSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillExperienceSheet(FileInfo excelFilePath, CRPlanning crp, List<PrestSante> myPrestData)
        {
            try
            {                
                CollectPrestaData(excelFilePath, crp, myPrestData, C.eExcelSheetPrestaData.Experience);
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillExperienceSheet : " + ex.Message);
                throw ex;
            }
        }

        //this is a temporary solution => get data from table _TempExpData and send them to excel
        public static void FillExperienceSheet2(FileInfo excelFilePath, DateTime debutPeriod, DateTime finPeriod)
        {
            try
            {
                //List<C_TempExpData> expData = new List<C_TempExpData>();
                //expData = C_TempExpData.GetExpData(debutPeriod.Year);

                List<C_TempExpData> expData = C_TempExpData.GetExpData(debutPeriod.Year, finPeriod.Year);

                var expDataWithoutId = expData.Select(e => new
                {
                    Au = e.Au,
                    Contrat = e.Contrat.Trim(),
                    Codcol = e.CodCol.Trim(),
                    AnneeExp = e.AnneeExp,
                    Libacte = e.LibActe.Trim(),
                    Libfam = e.LibFam.Trim(),
                    TypeCas = e.TypeCas.Trim(),
                    NbrActe = e.NombreActe,
                    FR = e.Fraisreel,
                    RSS = e.Rembss,
                    Rannexe = e.RembAnnexe,
                    Rnous = e.RembNous,
                    Res = e.Reseau.Trim(),
                    Minfr = e.MinFr,
                    Maxfr = e.MaxFr,
                    Minnous = e.MinNous,
                    Maxnous = e.MaxNous
                });

                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELEXP].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELEXP];

                    ws.Cells["A2"].LoadFromCollection(expDataWithoutId);

                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillExperienceSheet2 : " + ex.Message);
                throw ex;
            }
        }

        public static void FillProvisionSheet(FileInfo excelFilePath, CRPlanning crp, List<PrestSante> myPrestData)
        {
            try
            {
                CollectPrestaData(excelFilePath, crp, myPrestData, C.eExcelSheetPrestaData.Provision);
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillProvisionSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillQuartileSheet(FileInfo excelFilePath, List<PrestSante> myPrestData)
        {
            try
            {
                //create the table that holds the values for the quartiles
                DataTable quartileTable = new DataTable();

                DataColumn min = new DataColumn("MIN", typeof(decimal));
                DataColumn max = new DataColumn("MAX", typeof(decimal));
                DataColumn q1 = new DataColumn("Q1", typeof(decimal));
                DataColumn q2 = new DataColumn("Q2", typeof(decimal));
                DataColumn q3 = new DataColumn("Q3", typeof(decimal));
                DataColumn q4 = new DataColumn("Q4", typeof(decimal));
                DataColumn avg = new DataColumn("AVG", typeof(decimal));

                quartileTable.Columns.AddRange(new DataColumn[] { min, max, q1, q2, q3, q4, avg });

                //get garanty names to be treated
                List<string> garantyList = new List<string>();
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELQUARTILE];

                    //### get only values for Garanty Names: Second row, first column
                    for (int row = 2; row <= ws.Dimension.End.Row; row++)
                    {
                        if(ws.Cells[row, 3].Value != null)
                            garantyList.Add(ws.Cells[row, 3].Value.ToString());
                    }
                }

                foreach (string gar in garantyList)
                {
                    DataRow newRow = quartileTable.NewRow();

                    List<double> fraisReelList;

                    //the first case is no longer needed
                    if (gar.ToUpper().Contains("PROTHESES DENTAIRES ACCEPTEES  --- no longer needed"))
                    {
                       fraisReelList = myPrestData.Where(p => p.GarantyName.ToString().ToLower() == gar.ToLower() && p.FraisReel > 0 && p.PrixUnit == 107.5).
                       OrderBy(p => p.FraisReel.Value).
                       Select(p => p.FraisReel.Value / p.NombreActe.Value).ToList();
                    }
                    else
                    {
                       fraisReelList = myPrestData.Where(p => p.GarantyName.ToString().ToLower() == gar.ToLower() && p.FraisReel > 0 ).
                       OrderBy(p => p.FraisReel.Value).
                       Select(p => p.FraisReel.Value / p.NombreActe.Value).ToList();
                    }                    

                    int totalElements = fraisReelList.Count();
                    if (totalElements > 0)
                    {
                        double maxVal = fraisReelList.Max();
                        double minVal = fraisReelList.Min();
                        double avgVal = fraisReelList.Average();

                        int posQ1 = totalElements * C.cQuartile1 / 100;
                        int posQ2 = totalElements * C.cQuartile2 / 100;
                        int posQ3 = totalElements * C.cQuartile3 / 100;
                        int posQ4 = totalElements * C.cQuartile4 / 100;
                        
                        newRow["MIN"] = minVal;
                        newRow["MAX"] = maxVal;
                        newRow["Q1"] = totalElements > posQ1 ? fraisReelList[posQ1] : 0;
                        newRow["Q2"] = totalElements > posQ1 ? fraisReelList[posQ2] : 0;
                        newRow["Q3"] = totalElements > posQ1 ? fraisReelList[posQ3] : 0;
                        newRow["Q4"] = totalElements > posQ1 ? fraisReelList[posQ4] : 0;
                        newRow["AVG"] = avgVal;

                        quartileTable.Rows.Add(newRow);
                    }
                    else
                    {
                        //throw new Exception("THe Excel sheet for 'Quartiles' cannot be created, because no values were found for the following garanty: " + gar );

                        newRow["MIN"] = 0;
                        newRow["MAX"] = 0;
                        newRow["Q1"] = 0;
                        newRow["Q2"] = 0;
                        newRow["Q3"] = 0;
                        newRow["Q4"] = 0;
                        newRow["AVG"] = 0;

                        quartileTable.Rows.Add(newRow);
                    }
                }

                //save data to Excel
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    //pck.Workbook.Worksheets[C.cEXCELQUARTILE].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELQUARTILE];
                    ws.Cells["D2"].LoadFromDataTable(quartileTable, false);
                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillQuartileSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillAffichageSheet(FileInfo excelFilePath, string assur)
        {
            try
            {
                //create the table that holds the values for the quartiles
                DataTable affTable = new DataTable();
                DataTable affTable2 = new DataTable();

                //DataColumn order = new DataColumn("ORDER", typeof(int));
                DataColumn assureur = new DataColumn("ASSUREUR", typeof(string));
                DataColumn group = new DataColumn("GROUP", typeof(string));
                DataColumn garanty = new DataColumn("GARANTY", typeof(string));
                DataColumn assureur2 = new DataColumn("ASSUREUR", typeof(string));
                DataColumn group2 = new DataColumn("GROUP", typeof(string));

                affTable.Columns.AddRange(new DataColumn[] { assureur, group, garanty });
                affTable2.Columns.AddRange(new DataColumn[] { assureur2, group2 });

                var uniqueGGList = GroupGarantySante.GetUniqueGroupsAndGarantiesForAssureur(assur);
                var uniqueAGList = GroupGarantySante.GetUniqueAssureurAndGroups(assur);

                if (uniqueGGList.Any())
                {
                    foreach (var elem in uniqueGGList)
                    {
                        DataRow newRow = affTable.NewRow();

                        newRow["ASSUREUR"] = elem.AssureurName;
                        newRow["GROUP"] = elem.GroupName;
                        newRow["GARANTY"] = elem.GarantyName;
                        
                        affTable.Rows.Add(newRow);
                    }
                }

                if (uniqueAGList.Any())
                {
                    foreach (var elem in uniqueAGList)
                    {
                        DataRow newRow2 = affTable2.NewRow();

                        newRow2["ASSUREUR"] = elem.AssureurName;
                        newRow2["GROUP"] = elem.GroupName;
                        
                        affTable2.Rows.Add(newRow2);
                    }
                }

                //save data to Excel
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELGROUPGARANT].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELGROUPGARANT];
                    ws.Cells["A2"].LoadFromDataTable(affTable, false);
                    ws.Cells["E2"].LoadFromDataTable(affTable2, false);
                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillPrevSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillDates(FileInfo excelFilePath, DateTime dateArret, DateTime debutPeriode, DateTime finPeriode,
            double? TaxDef, double? TaxAct, double? TaxPer, bool? calculProvision)
        {
            try
            {
                //save data to Excel
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {                   
                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELGROUPGARANT];

                    if (calculProvision.HasValue)
                    {
                        ws.Cells["A2"].Value = calculProvision.Value == true ? "OUI" : "NON";
                    }

                    //if (TaxDef.HasValue)
                    //    ws.Cells["Q2"].Value = TaxDef/100; 
                    //if (TaxAct.HasValue)
                    //    ws.Cells["R2"].Value = TaxAct/100;
                    //if (TaxPer.HasValue)
                    //    ws.Cells["S2"].Value = TaxPer/100;

                    ws.Cells["M2"].Value = debutPeriode.ToShortDateString();
                    ws.Cells["N2"].Value = finPeriode.ToShortDateString();
                    ws.Cells["O2"].Value = dateArret.ToShortDateString();

                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillAffichageSheet : " + ex.Message);
                throw ex;
            }
        }

        public static void FillOUI(FileInfo excelFilePath)
        {
            try
            {
                //save data to Excel
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELPAGEGARDE];
                    
                    ws.Cells["AD2"].Value = "OUI";
                    //ws.Cells["AE2"].Value = "OUI";

                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: FillAffichageSheet : " + ex.Message);
                throw ex;
            }
        }


        //### this is almost a duplicate of the Method: CollectPrestaData
        // this needs to be cleaned up and improved
        private static void CollectPrestaData2(FileInfo excelFilePath, CRPlanning crp, List<PrestSante> myPrestData, C.eExcelSheetPrestaData excelSheet)
        {
            try
            {
                string myExcelSheet = C.cEXCELPRESTWITHOPTION;
                List<Cadencier> cadencierAll = new List<Cadencier>();               

                DataTable prestaTable = CreatePrestaPrevExpTable(myExcelSheet);

                //CodeCol = p.Select(pr=>pr.CodeCol).First(),
                List<ExcelPrestaSheet> excelPrestDataSmall = myPrestData
                           .GroupBy(p => new
                           {
                               DateSoinsYear = p.DateSoins.Value.Year,
                               p.GroupName,
                               p.GarantyName,
                               CAS2 = p.CAS.ToLower() == "true" ? "VRAI" : "FAUX",
                               RES = String.IsNullOrEmpty(p.Reseau) ? "FAUX" : "VRAI"
                           })
                           .Select(p => new ExcelPrestaSheet
                           {
                               DateVision = new DateTime(1900, 01, 01),
                               ContractId = "XXXXX",
                               CodeCol = "XXXXX", // p.Select(pr=>pr.CodeCol).First(),
                               DateSoins = new DateTime(p.Key.DateSoinsYear, 1, 1),
                               GroupName = p.Key.GroupName,
                               GarantyName = p.Key.GarantyName,
                               CAS = p.Key.CAS2,
                               NombreActe = p.Sum(pr => pr.NombreActe),  //.Where(pr => pr.NombreActe >= 0)
                               FraisReel = p.Sum(pr => pr.FraisReel),
                               RembSS = p.Sum(pr => pr.RembSS),
                               RembAnnexe = p.Sum(pr => pr.RembAnnexe),
                               RembNous = p.Sum(pr => pr.RembNous),
                               Reseau = p.Key.RES,
                               MinFR = p.Where(pr => pr.FraisReel >= 0).Min(pr => pr.FraisReel / pr.NombreActe),
                               MaxFR = p.Where(pr => pr.FraisReel >= 0).Max(pr => pr.FraisReel / pr.NombreActe),
                               MinNous = p.Where(pr => pr.RembNous >= 0).Min(pr => pr.RembNous / pr.NombreActe),
                               MaxNous = p.Where(pr => pr.RembNous >= 0).Max(pr => pr.RembNous / pr.NombreActe)                               
                           })
                           //.Where(p => p.GarantyName == "LENTILLES")
                           .OrderBy(gr => gr.GroupName).ThenBy(ga => ga.GarantyName)
                           .ToList();

                List<ExcelPrestaSheet> excelPrestDataLarge = myPrestData
                           .GroupBy(p => new
                           {
                               p.DateVision,
                               p.ContractId,
                               p.CodeCol,
                               DateSoinsYear = p.DateSoins.Value.Year,
                               p.GroupName,
                               p.GarantyName,
                               CAS2 = p.CAS.ToLower() == "true" ? "VRAI" : "FAUX",
                               RES = String.IsNullOrEmpty(p.Reseau) ? "FAUX" : "VRAI",
                               p.BO1,
                               p.BO2
                           })
                           .Select(p => new ExcelPrestaSheet
                           {
                               DateVision = p.Key.DateVision,
                               ContractId = p.Key.ContractId,
                               CodeCol = p.Key.CodeCol,
                               DateSoins = new DateTime(p.Key.DateSoinsYear, 1, 1),
                               GroupName = p.Key.GroupName,
                               GarantyName = p.Key.GarantyName,
                               CAS = p.Key.CAS2,
                               NombreActe = p.Sum(pr => pr.NombreActe),  //.Where(pr => pr.NombreActe >= 0)
                               FraisReel = p.Sum(pr => pr.FraisReel),
                               RembSS = p.Sum(pr => pr.RembSS),
                               RembAnnexe = p.Sum(pr => pr.RembAnnexe),
                               RembNous = p.Sum(pr => pr.RembNous),
                               Reseau = p.Key.RES,
                               MinFR = p.Where(pr => pr.FraisReel >= 0).Min(pr => pr.FraisReel / pr.NombreActe),
                               MaxFR = p.Where(pr => pr.FraisReel >= 0).Max(pr => pr.FraisReel / pr.NombreActe),
                               MinNous = p.Where(pr => pr.RembNous >= 0).Min(pr => pr.RembNous / pr.NombreActe),
                               MaxNous = p.Where(pr => pr.RembNous >= 0).Max(pr => pr.RembNous / pr.NombreActe),
                               BO1 = p.Key.BO1,
                               BO2 = p.Key.BO2
                           })
                           //.Where(p => p.GarantyName == "LENTILLES")
                           .OrderBy(gr => gr.GroupName).ThenBy(ga => ga.GarantyName)
                           .ToList();


                //### create 2 lists => iterate through large list (with college) and replace all calc fields (min, max...) with calc fields from 
                // corresponding line in small list (use 5 key fields)
                // var item = smallList.FirstOrDefault(o => o.GroupName == groupName && ...);
                //if (item != null)
                //    item.value = "Value";

                foreach (ExcelPrestaSheet dat in excelPrestDataLarge)
                {
                    var item = excelPrestDataSmall.FirstOrDefault(i => i.DateSoins == dat.DateSoins && i.GroupName == dat.GroupName && i.GarantyName == dat.GarantyName
                        && i.CAS == dat.CAS && i.Reseau == dat.Reseau);

                    if (item != null)
                    {
                        dat.MinFR = item.MinFR;
                        dat.MaxFR = item.MaxFR;
                        dat.MinNous = item.MinNous;
                        dat.MaxNous = item.MaxNous;

                    }
                }

                //foreach (PrestSante prest in myPrestData)
                foreach (ExcelPrestaSheet prest in excelPrestDataLarge)
                {                    
                    DataRow newRow = prestaTable.NewRow();

                    newRow["ANNEESOIN"] = prest.DateSoins.HasValue ? prest.DateSoins.Value.Year : 0;
                    newRow["AU"] = prest.DateVision.HasValue ? prest.DateVision.Value : (object)DBNull.Value;
                    newRow["CONTRAT"] = prest.ContractId;
                    newRow["CODCOL"] = prest.CodeCol;
                    newRow["LIBACTE"] = prest.GarantyName;
                    newRow["LIBFAM"] = prest.GroupName;
                    
                    newRow["NBREACTE"] = prest.NombreActe.HasValue ? prest.NombreActe : 0;
                    newRow["FRAISREELS"] = prest.FraisReel.HasValue ? prest.FraisReel.Value : 0;
                    newRow["REMBSS"] = prest.RembSS.HasValue ? prest.RembSS.Value : 0;
                    newRow["REMBANNEXE"] = prest.RembAnnexe.HasValue ? prest.RembAnnexe.Value : 0;
                    newRow["REMBNOUS"] = prest.RembNous.HasValue ? prest.RembNous.Value : 0;
                    newRow["CASNONCAS"] = prest.CAS;
                    newRow["RESEAU"] = prest.Reseau;
                    newRow["MINFR"] = prest.MinFR.HasValue ? prest.MinFR.Value : 0;
                    newRow["MAXFR"] = prest.MaxFR.HasValue ? prest.MaxFR.Value : 0;
                    newRow["MINNOUS"] = prest.MinNous.HasValue ? prest.MinNous.Value : 0;
                    newRow["MAXNOUS"] = prest.MaxNous.HasValue ? prest.MaxNous.Value : 0;

                    newRow["BO1"] = prest.BO1;
                    newRow["BO2"] = prest.BO2;


                    prestaTable.Rows.Add(newRow);
                }

                //save to Excel
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[C.cEXCELPREST].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[C.cEXCELPREST];
                    ws.Cells["A2"].LoadFromDataTable(prestaTable, false);
                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: CollectPrestaData2 : " + ex.Message);
                throw ex;
            }
        }
        

        public static List<ExcelPrestaSheet> GenerateModifiedPrestData(List<PrestSante> myPrestData)
        {
            try
            {
                //CodeCol = p.Select(pr=>pr.CodeCol).First(),
                List<ExcelPrestaSheet> excelPrestDataSmall = myPrestData
                           .GroupBy(p => new
                           {
                               DateSoinsYear = p.DateSoins.Value.Year,
                               p.GroupName,
                               p.GarantyName,
                               CAS2 = p.CAS.ToLower() == "true" ? "VRAI" : "FAUX",
                               RES = String.IsNullOrEmpty(p.Reseau) ? "FAUX" : "VRAI"
                           })
                           .Select(p => new ExcelPrestaSheet
                           {
                               DateVision = new DateTime(1900, 01, 01),
                               ContractId = "XXXXX",
                               CodeCol = "XXXXX", // p.Select(pr=>pr.CodeCol).First(),
                               DateSoins = new DateTime(p.Key.DateSoinsYear, 1, 1),
                               GroupName = p.Key.GroupName,
                               GarantyName = p.Key.GarantyName,
                               CAS = p.Key.CAS2,
                               NombreActe = p.Sum(pr => pr.NombreActe),  //.Where(pr => pr.NombreActe >= 0)
                               FraisReel = p.Sum(pr => pr.FraisReel),
                               RembSS = p.Sum(pr => pr.RembSS),
                               RembAnnexe = p.Sum(pr => pr.RembAnnexe),
                               RembNous = p.Sum(pr => pr.RembNous),
                               Reseau = p.Key.RES,
                               MinFR = p.Where(pr => pr.FraisReel >= 0).Min(pr => pr.FraisReel / pr.NombreActe),
                               MaxFR = p.Where(pr => pr.FraisReel >= 0).Max(pr => pr.FraisReel / pr.NombreActe),
                               MinNous = p.Where(pr => pr.RembNous >= 0).Min(pr => pr.RembNous / pr.NombreActe),
                               MaxNous = p.Where(pr => pr.RembNous >= 0).Max(pr => pr.RembNous / pr.NombreActe)
                           })
                           //.Where(p => p.GarantyName == "LENTILLES")
                           .OrderBy(gr => gr.GroupName).ThenBy(ga => ga.GarantyName)
                           .ToList();

                List<ExcelPrestaSheet> excelPrestDataLarge = myPrestData
                           .GroupBy(p => new
                           {
                               p.DateVision,
                               p.ContractId,
                               p.CodeCol,
                               DateSoinsYear = p.DateSoins.Value.Year,
                               p.GroupName,
                               p.GarantyName,
                               CAS2 = p.CAS.ToLower() == "true" ? "VRAI" : "FAUX",
                               RES = String.IsNullOrEmpty(p.Reseau) ? "FAUX" : "VRAI"
                           })
                           .Select(p => new ExcelPrestaSheet
                           {
                               DateVision = p.Key.DateVision,
                               ContractId = p.Key.ContractId,
                               CodeCol = p.Key.CodeCol,
                               DateSoins = new DateTime(p.Key.DateSoinsYear, 1, 1),
                               GroupName = p.Key.GroupName,
                               GarantyName = p.Key.GarantyName,
                               CAS = p.Key.CAS2,
                               NombreActe = p.Sum(pr => pr.NombreActe),  //.Where(pr => pr.NombreActe >= 0)
                               FraisReel = p.Sum(pr => pr.FraisReel),
                               RembSS = p.Sum(pr => pr.RembSS),
                               RembAnnexe = p.Sum(pr => pr.RembAnnexe),
                               RembNous = p.Sum(pr => pr.RembNous),
                               Reseau = p.Key.RES,
                               MinFR = p.Where(pr => pr.FraisReel >= 0).Min(pr => pr.FraisReel / pr.NombreActe),
                               MaxFR = p.Where(pr => pr.FraisReel >= 0).Max(pr => pr.FraisReel / pr.NombreActe),
                               MinNous = p.Where(pr => pr.RembNous >= 0).Min(pr => pr.RembNous / pr.NombreActe),
                               MaxNous = p.Where(pr => pr.RembNous >= 0).Max(pr => pr.RembNous / pr.NombreActe)
                           })
                           //.Where(p => p.GarantyName == "LENTILLES")
                           .OrderBy(gr => gr.GroupName).ThenBy(ga => ga.GarantyName)
                           .ToList();


                //### create 2 lists => iterate through large list (with college) and replace all calc fields (min, max...) with calc fields from 
                // corresponding line in small list (use 5 key fields)
                // var item = smallList.FirstOrDefault(o => o.GroupName == groupName && ...);
                //if (item != null)
                //    item.value = "Value";

                foreach (ExcelPrestaSheet dat in excelPrestDataLarge)
                {
                    var item = excelPrestDataSmall.FirstOrDefault(i => i.DateSoins == dat.DateSoins && i.GroupName == dat.GroupName && i.GarantyName == dat.GarantyName
                        && i.CAS == dat.CAS && i.Reseau == dat.Reseau);

                    if (item != null)
                    {
                        dat.MinFR = item.MinFR;
                        dat.MaxFR = item.MaxFR;
                        dat.MinNous = item.MinNous;
                        dat.MaxNous = item.MaxNous;

                    }
                }

                return excelPrestDataLarge;
            }
            catch (Exception ex)
            {
                log.Error("Error :: GenerateModifiedPrestData : " + ex.Message);
                throw ex;
            }

        }
                 

        private static void CollectPrestaData(FileInfo excelFilePath, CRPlanning crp, List<PrestSante> myPrestData, C.eExcelSheetPrestaData excelSheet)
        {
            try
            {
                string myExcelSheet = C.cEXCELPREST;                
                List<Cadencier> cadencierAll = new List<Cadencier>();                

                switch (excelSheet)
                {
                    case C.eExcelSheetPrestaData.Prestation:
                        myExcelSheet = C.cEXCELPREST;
                        break;
                    case C.eExcelSheetPrestaData.Experience:
                        myExcelSheet = C.cEXCELEXP;
                        break;
                    case C.eExcelSheetPrestaData.Provision:
                        myExcelSheet = C.cEXCELPROV;
                        break;
                    default:
                        myExcelSheet = C.cEXCELPREST;
                        break;
                }

                if (myExcelSheet == C.cEXCELPROV)
                {
                    //myCad = Cadencier.GetCadencierForAssureurId(25);

                    List<string> assList = myPrestData.OrderBy(p => p.AssureurName).Select(p => p.AssureurName).Distinct().ToList();
                    List<Cadencier> cadencierForAssureur = new List<Cadencier>();
                    cadencierAll = Cadencier.GetCadencierForAssureur(C.cDEFAULTASSUREUR);

                    foreach (string assurName in assList)
                    {
                        if (assurName != C.cDEFAULTASSUREUR)
                        {
                            cadencierForAssureur = Cadencier.GetCadencierForAssureur(assurName);
                            cadencierAll.AddRange(cadencierForAssureur);
                        }
                    }
                }                

               // string mycsv = CsvSerializer.SerializeToCsv(myPrestData);

                DataTable prestaTable = CreatePrestaPrevExpTable(myExcelSheet);

                //### get data from external function
                List<ExcelPrestaSheet> excelPrestDataLarge = GenerateModifiedPrestData(myPrestData);

                //foreach (PrestSante prest in myPrestData)
                foreach (ExcelPrestaSheet prest in excelPrestDataLarge)
                {
                    //GroupGarantyPair ggPair = GetGroupGarantyPairForCodeActe(groupSanteListForAssureur, prest.CodeActe);                    

                    DataRow newRow = prestaTable.NewRow();
                    
                    newRow["ANNEESOIN"] = prest.DateSoins.HasValue ? prest.DateSoins.Value.Year : 0;

                    if (excelSheet == C.eExcelSheetPrestaData.Experience)
                    {
                        int expYear = int.Parse(newRow["ANNEESOIN"].ToString());
                        newRow["ANNEESOIN"] = expYear; // + 1;
                    }


                    newRow["AU"] = prest.DateVision.HasValue ? prest.DateVision.Value : (object)DBNull.Value; 
                    newRow["CONTRAT"] = prest.ContractId;
                    newRow["CODCOL"] = prest.CodeCol;
                    newRow["LIBACTE"] = prest.GarantyName;
                    newRow["LIBFAM"] = prest.GroupName;

                    if (myExcelSheet == C.cEXCELPROV)
                    {                        
                        double rembNous = prest.RembNous.HasValue ? prest.RembNous.Value : 0;
                        int anneeSoins = prest.DateSoins.HasValue ? prest.DateSoins.Value.Year : 0;
                        DateTime dateArret = crp.DateArret.HasValue ? crp.DateArret.Value : DateTime.MinValue;
                        DateTime dateDebutPeriode = crp.DebutPeriode.HasValue ? crp.DebutPeriode.Value : DateTime.MinValue;
                        DateTime dateFinPeriode = crp.FinPeriode.HasValue ? crp.FinPeriode.Value : DateTime.MinValue;

                        //#### modify this: provide Assureur
                        double coefCad = GetCoefCadencier(anneeSoins, dateArret, dateDebutPeriode, dateFinPeriode, cadencierAll, null);
                        newRow["PROVISION"] = coefCad * rembNous;
                    }
                    else
                    {
                        newRow["NBREACTE"] = prest.NombreActe.HasValue ? prest.NombreActe : 0;
                        newRow["FRAISREELS"] = prest.FraisReel.HasValue ? prest.FraisReel.Value : 0;
                        newRow["REMBSS"] = prest.RembSS.HasValue ? prest.RembSS.Value : 0;
                        newRow["REMBANNEXE"] = prest.RembAnnexe.HasValue ? prest.RembAnnexe.Value : 0;
                        newRow["REMBNOUS"] = prest.RembNous.HasValue ? prest.RembNous.Value : 0;
                        newRow["CASNONCAS"] = prest.CAS;
                        newRow["RESEAU"] = prest.Reseau;
                        newRow["MINFR"] = prest.MinFR.HasValue ? prest.MinFR.Value : 0;
                        newRow["MAXFR"] = prest.MaxFR.HasValue ? prest.MaxFR.Value : 0;
                        newRow["MINNOUS"] = prest.MinNous.HasValue ? prest.MinNous.Value : 0;
                        newRow["MAXNOUS"] = prest.MaxNous.HasValue ? prest.MaxNous.Value : 0;
                    }

                    prestaTable.Rows.Add(newRow);
                }

                //save to Excel
                using (ExcelPackage pck = new ExcelPackage(excelFilePath))
                {
                    pck.Workbook.Worksheets[myExcelSheet].DeleteRow(2, C.cNUMBROWSDELETEEXCEL);

                    ExcelWorksheet ws = pck.Workbook.Worksheets[myExcelSheet];
                    ws.Cells["A2"].LoadFromDataTable(prestaTable, false);
                    pck.Save();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error :: CollectPrestaData : " + ex.Message);
                throw ex;
            }
        }

        private static DataTable CreatePrestaPrevExpTable(string myExcelSheet)
        {
            try
            {
                DataTable myTable = new DataTable();

                DataColumn au = new DataColumn("AU", typeof(DateTime));
                DataColumn contrat = new DataColumn("CONTRAT", typeof(string));
                DataColumn codcol = new DataColumn("CODCOL", typeof(string));
                DataColumn anneesoins = new DataColumn("ANNEESOIN", typeof(int));
                DataColumn libacte = new DataColumn("LIBACTE", typeof(string));
                DataColumn libfam = new DataColumn("LIBFAM", typeof(string));
                DataColumn casNonCas = new DataColumn("CASNONCAS", typeof(string));
                DataColumn nbreacte = new DataColumn("NBREACTE", typeof(int));
                DataColumn fraisreels = new DataColumn("FRAISREELS", typeof(decimal));
                DataColumn rembss = new DataColumn("REMBSS", typeof(decimal));
                DataColumn rembannexe = new DataColumn("REMBANNEXE", typeof(decimal));
                DataColumn rembnous = new DataColumn("REMBNOUS", typeof(decimal));
                DataColumn provision = new DataColumn("PROVISION", typeof(decimal));
                DataColumn reseau = new DataColumn("RESEAU", typeof(string));
                DataColumn minfr = new DataColumn("MINFR", typeof(decimal));
                DataColumn maxfr = new DataColumn("MAXFR", typeof(decimal));
                DataColumn minnous = new DataColumn("MINNOUS", typeof(decimal));
                DataColumn maxnous = new DataColumn("MAXNOUS", typeof(decimal));
                DataColumn famcoll = new DataColumn("FAMILLECOLLEGE", typeof(string));
                DataColumn bo1 = new DataColumn("BO1", typeof(string));
                DataColumn bo2 = new DataColumn("BO2", typeof(string));

                if (myExcelSheet == C.cEXCELPROV)
                    myTable.Columns.AddRange(new DataColumn[] { au, contrat, codcol, anneesoins, libacte, libfam, provision });
                else if (myExcelSheet == C.cEXCELEXP)
                    myTable.Columns.AddRange(new DataColumn[] { au, contrat, codcol, anneesoins, libacte, libfam, casNonCas,
                    nbreacte, fraisreels, rembss, rembannexe, rembnous, reseau, minfr, maxfr, minnous, maxnous});
                else if (myExcelSheet == C.cEXCELPREST)
                    myTable.Columns.AddRange(new DataColumn[] { au, contrat, codcol, anneesoins, libacte, libfam, casNonCas,
                    nbreacte, fraisreels, rembss, rembannexe, rembnous, reseau, minfr, maxfr, minnous, maxnous});
                else if (myExcelSheet == C.cEXCELPRESTWITHOPTION)
                    myTable.Columns.AddRange(new DataColumn[] { au, contrat, codcol, anneesoins, libacte, libfam, casNonCas,
                    nbreacte, fraisreels, rembss, rembannexe, rembnous, reseau, minfr, maxfr, minnous, maxnous, famcoll, bo1, bo2});

                return myTable;
            }
            catch (Exception ex)
            {
                log.Error("Error :: CreatePrestaPrevExpTable : " + ex.Message);
                throw ex;
            }

        }

        private static DataTable CreateGlobalTable(C.eReportTypes reportType)
        {
            try
            {
                DataTable myTable = new DataTable();

                DataColumn Assureur = new DataColumn("Assureur", typeof(string));
                DataColumn Company = new DataColumn("Company", typeof(string));
                DataColumn Subsid = new DataColumn("Subsid", typeof(string));
                DataColumn YearSurv = new DataColumn("YearSurv", typeof(int));
                DataColumn RNous = new DataColumn("RNous", typeof(decimal));
                DataColumn Provisions = new DataColumn("Provisions", typeof(decimal));
                DataColumn CotBrut = new DataColumn("CotBrut", typeof(decimal));
                //DataColumn TaxTotal = new DataColumn("TaxTotal", typeof(string));
                DataColumn CotNet = new DataColumn("CotNet", typeof(decimal));
                DataColumn Ratio = new DataColumn("Ratio", typeof(decimal));
                DataColumn GainLoss = new DataColumn("GainLoss", typeof(decimal));
                DataColumn CoeffCad = new DataColumn("CoeffCad", typeof(decimal));
                DataColumn FR = new DataColumn("FR", typeof(decimal));
                DataColumn RSS = new DataColumn("RSS", typeof(decimal));
                DataColumn RAnnexe = new DataColumn("RAnnexe", typeof(decimal));                
                //DataColumn TaxDefault = new DataColumn("TaxDefault", typeof(string));
                //DataColumn TaxActive = new DataColumn("TaxActive", typeof(string));
                DataColumn DateArret = new DataColumn("DateArret", typeof(DateTime));

                //myTable.Columns.AddRange(new DataColumn[] { Assureur, Company, Subsid, YearSurv, FR, RSS, RAnnexe, RNous, Provisions, CoeffCad, CotBrut, TaxTotal, TaxDefault, TaxActive,
                //    CotNet, Ratio, GainLoss, DateArret });

                //with taxes
                //myTable.Columns.AddRange(new DataColumn[] { Assureur, Company, Subsid, YearSurv, RNous, Provisions, CotBrut, TaxTotal, CotNet, Ratio, GainLoss, CoeffCad,
                //    FR, RSS, RAnnexe, TaxDefault, TaxActive, DateArret });
                //without Taxes
                myTable.Columns.AddRange(new DataColumn[] { Assureur, Company, Subsid, YearSurv, RNous, Provisions, CotBrut, CotNet, Ratio, GainLoss, CoeffCad,
                    FR, RSS, RAnnexe, DateArret });

                //if (reportType == C.eReportTypes.GlobalEnt)
                //    myTable.Columns.AddRange(new DataColumn[] { Assureur, Company, YearSurv, FR, RSS, RAnnexe, RNous, Provisions, CoeffCad, CotBrut, TaxTotal, TaxDefault, TaxActive,
                //    CotNet, Ratio, GainLoss, DateArret });
                //else if (reportType == C.eReportTypes.GlobalSubsid)
                //    myTable.Columns.AddRange(new DataColumn[] { Assureur, Company, Subsid, YearSurv, FR, RSS, RAnnexe, RNous, Provisions, CoeffCad, CotBrut, TaxTotal, TaxDefault, TaxActive,
                //    CotNet, Ratio, GainLoss, DateArret });

                return myTable;
            }
            catch (Exception ex)
            {
                log.Error("Error :: CreatePrestaPrevExpTable : " + ex.Message);
                throw ex;
            }

        }

        private static DataTable CreateGlobalTablePrev(C.eReportTypes reportType)
        {
            try
            {
                DataTable myTable = new DataTable();

                DataColumn Assureur = new DataColumn("Assureur", typeof(string));
                DataColumn Company = new DataColumn("Company", typeof(string));
                DataColumn Subsid = new DataColumn("Subsid", typeof(string));
                DataColumn YearSurv = new DataColumn("YearSurv", typeof(int));
                DataColumn TypeSinistre = new DataColumn("TypeSinistre", typeof(string));
                DataColumn Prestations = new DataColumn("Prestations", typeof(decimal));
                DataColumn Provisions = new DataColumn("Provisions", typeof(decimal));
                DataColumn CotBrute = new DataColumn("CotBrute", typeof(decimal));
                DataColumn TauxChargement = new DataColumn("TauxChargement", typeof(string));
                DataColumn CotNet = new DataColumn("CotNet", typeof(decimal));
                DataColumn Ratio = new DataColumn("Ratio", typeof(decimal));
                DataColumn GainLoss = new DataColumn("GainLoss", typeof(decimal));
                DataColumn DateArret = new DataColumn("DateArret", typeof(DateTime));
                
                myTable.Columns.AddRange(new DataColumn[] { Assureur, Company, Subsid, YearSurv, TypeSinistre, Prestations, Provisions, CotBrute, TauxChargement, CotNet, Ratio, GainLoss, 
                    DateArret });

                return myTable;
            }
            catch (Exception ex)
            {
                log.Error("Error :: CreatePrestaPrevExpTable : " + ex.Message);
                throw ex;
            }

        }

        private static double GetCoefCadencier(int anneeSoins, DateTime dateArret, DateTime dateDebutPeriode,
            DateTime dateFinPeriode, List<Cadencier> myCad, string assur)
        {
            try
            {
                double cumul = 0;
                int month = 0;

                //double rembNous = prest.RembNous.HasValue ? prest.RembNous.Value : 0;
                //int anneeSoins = prest.DateSoins.HasValue ? prest.DateSoins.Value.Year : 0;

                DateTime date1;
                DateTime dateDebutPeriodeAdjusted;
                DateTime dateFinPeriodeAdjusted;

                if (anneeSoins != 0 && dateArret != DateTime.MinValue && dateFinPeriode != DateTime.MinValue)
                {
                    date1 = new DateTime(anneeSoins, dateDebutPeriode.Month, dateDebutPeriode.Day);
                    //month = ((dateArret.Year - date1.Year) * 12) + dateArret.Month - date1.Month;

                    TimeSpan span = dateArret.Subtract(date1);
                    double monthDouble = span.TotalDays / 30.25;
                    month = (int)Math.Round(monthDouble, MidpointRounding.AwayFromZero);

                    dateDebutPeriodeAdjusted = new DateTime(anneeSoins, dateDebutPeriode.Month, dateDebutPeriode.Day);
                    dateFinPeriodeAdjusted = new DateTime(anneeSoins, dateFinPeriode.Month, dateFinPeriode.Day);

                    var res = myCad.Where(c => c.Month == month && c.Year == anneeSoins && c.DebutSurvenance == dateDebutPeriodeAdjusted
                        && c.FinSurvenance == dateFinPeriodeAdjusted);
                    if (res.Any())
                    {
                        //choose the value according to the provided Assureur
                        if (string.IsNullOrWhiteSpace(assur))
                            cumul = res.ToList()[0].Cumul.Value;
                        else
                        {
                            var cadVal = res.Where(c => c.AssureurName.ToLower() == assur.ToLower()).First();
                            if (cadVal != null)
                                cumul = cadVal.Cumul.Value;
                            else
                                cumul = 0;
                        }
                    }
                    else
                    {
                        cumul = 0;
                    }
                }
                else
                {
                    //we have  aproblem => log an error message                            
                    throw new Exception("A value for 'Provision' cannot be calculated! Provided values: Annee Soins: " + anneeSoins
                            + " date fin péeriode: " + dateFinPeriode.ToShortDateString() + " date arret: " + dateArret.ToShortDateString());
                }

                return cumul;
            }
            catch (Exception ex)
            {
                log.Error("Error :: GetCoefCadencier : " + ex.Message);
                throw ex;
            }
        }        

        private static double CalculateProvision(double rembNous, int anneeSoins, DateTime dateArret, DateTime dateDebutPeriode,
            DateTime dateFinPeriode, List<Cadencier> myCad, string assur)
        {
            try
            {
                double cumul = GetCoefCadencier(anneeSoins, dateArret, dateDebutPeriode, dateFinPeriode, myCad, assur);
                return cumul * rembNous; 
            }
            catch (Exception ex)
            {
                log.Error("Error :: CalculateProvision : " + ex.Message);
                throw ex;
            }
        }
        

        #region OLD METHODS

        //private static GroupGarantyPair GetGroupGarantyPairForCodeActe(List<GroupSante> groupSanteListForCompany, string codeActe)
        //{
        //    try
        //    {
        //        GroupGarantyPair ggPair = new GroupGarantyPair();
        //        string garName = "";
        //        string groupName = "";

        //        List<GroupSante> myGroupSante = groupSanteListForCompany.Where(gr => gr.GarantySantes.Any(gar => gar.CodeActe == codeActe)).ToList();

        //        if (myGroupSante.Any())
        //        {
        //            //we simply take the first element
        //            groupName = myGroupSante[0].Name;

        //            if (myGroupSante[0].GarantySantes.Count > 0)
        //                garName = myGroupSante[0].GarantySantes.ToList()[0].Name;
        //            else
        //                garName = "";
        //        }
        //        else
        //        {
        //            garName = "";
        //            groupName = "";
        //        }

        //        ggPair.GroupName = groupName;
        //        ggPair.GarantyName = garName;

        //        return ggPair;
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex.Message);
        //        throw ex;
        //    }
        //}


        #endregion

        
    }


}
