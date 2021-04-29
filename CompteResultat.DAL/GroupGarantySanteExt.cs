using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity.Validation;
using CompteResultat.Common;

namespace CompteResultat.DAL
{
    [MetadataType(typeof(GroupGarantySante.MetaData))]
    public partial class GroupGarantySante
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public static List<GroupGarantySante> GetGroupsAndGarantiesForAssureur(string assurName)
        {
            try
            {
                List<GroupGarantySante> ggs;

                using (var context = new CompteResultatEntities())
                {
                    ggs = context.GroupGarantySantes.Where(c => c.AssureurName == assurName).ToList();                    
                }

                //if (ggs == null || ggs.Count == 0)
                //    throw new Exception("The 'GroupGarantySante' entity does not contain any data!");

                return ggs;

            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }
        
        public static void DeleteGroupsWithSpecificAssureurName(string assurName)
        {
            try
            {
                using (var context = new CompteResultatEntities())
                {
                    context.GroupGarantySantes.RemoveRange(context.GroupGarantySantes.Where(g => g.AssureurName == assurName));
                    context.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static void TruncateTable()
        {
            try
            {
                using (var context = new CompteResultatEntities())
                {
                    //context.Database.ExecuteSqlCommand("TRUNCATE TABLE GroupGarantySante;");
                    context.Database.ExecuteSqlCommand("Delete From GroupGarantySante Where AssureurName Not Like 'Paramètres par défaut'");
                }
            }
            catch (Exception ex)
            {
                log.Error(ex.Message);
                throw ex;
            }
        }

        public static int InsertGroupGaranty(GroupGarantySante ggs)
        {
            try
            {
                using (var context = new CompteResultatEntities())
                {
                    //group.Id = GenerateCompanyID();

                    context.GroupGarantySantes.Add(ggs);
                    context.SaveChanges();

                    return ggs.Id;
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