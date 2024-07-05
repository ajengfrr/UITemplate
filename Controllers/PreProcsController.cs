using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using System.Data;
using Microsoft.CodeAnalysis;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using PreProc.Models;
using Microsoft.Data.SqlClient;
using System.Net.Mail;
using System.Numerics;
using Microsoft.IdentityModel.Tokens;
using System.Globalization;
using System.Reflection.Emit;
using System.Linq;
using System.Runtime.Intrinsics.Arm;

namespace PreProcsController.Controllers
{
    [Authorize]
    public class PreProcsController : Controller
    {
        private readonly PreProcContext _context;
        private readonly IConfiguration _configuration;

        //private readonly GraphServiceClient _graphClient;
        public PreProcsController(PreProcContext context, IConfiguration configuration)//, PreProcContextProcedures dbsp)
        {
            _context = context;
            _configuration = configuration;
            //_graphClient = graphClient;
        }

        public string getUserNow()
        {
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "10.10.62.17";
                builder.UserID = "sa";
                builder.Password = "Password1!";
                builder.InitialCatalog = "eRequisition";
                builder.TrustServerCertificate = true;
                builder.MultipleActiveResultSets = true;
                string multipolar = @"MULTIPOLAR\";
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    String sql = "SELECT * FROM gp_ms_Emp WHERE [EmailAddress] = '" + User.Identity.Name.Replace(multipolar, "") + "@multipolar.com'";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ViewBag.usernow = reader["NAMA"].ToString();
                                ViewBag.usernowemail = reader["EmailAddress"].ToString();
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (SqlException e)
            {

            }
            return ViewBag.usernow;
        }

        public string getUserNowEmail()
        {
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "10.10.62.17";
                builder.UserID = "sa";
                builder.Password = "Password1!";
                builder.InitialCatalog = "eRequisition";
                builder.TrustServerCertificate = true;
                builder.MultipleActiveResultSets = true;
                string multipolar = @"MULTIPOLAR\";
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    String sql = "SELECT * FROM gp_ms_Emp WHERE [EmailAddress] = '" + User.Identity.Name.Replace(multipolar, "") + "@multipolar.com'";
                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                ViewBag.usernowemail = reader["EmailAddress"].ToString();
                            }
                        }
                    }
                    connection.Close();
                }
            }
            catch (SqlException e)
            {

            }
            return ViewBag.usernowemail;
        }

        public async Task<IActionResult> ConfidentialDataAsync()
        {
            return View();
        }
        //
        //public async Task<IActionResult> Index()
        //{
        //    var preproc = from t in _context.PreProcGeneralInfo2s
        //                  select t;
        //    //var preproc = await _context.PreProcGeneralInfo2s.Join(_context.VwEmployeeBaseEreqs,
        //    //            p => p.Am, d => d.EmailAddress,
        //    //            (riskdata, risk) => new
        //    //            {
        //    //                PreProcId = riskdata.PreProcId
        //    //              ,
        //    //                RiskOrder = riskdata.RiskOrder
        //    //              ,
        //    //                CheckedBy = riskdata.CheckedBy
        //    //              ,
        //    //                RiskType = riskdata.RiskType
        //    //              ,
        //    //                RiskId = riskdata.RiskId
        //    //                ,
        //    //                RiskName = risk.RiskName
        //    //              ,
        //    //                RiskWeight = risk.RiskWeight
        //    //              ,
        //    //                IsActive = risk.IsActive
        //    //            }).Where(x => x.IsActive == true).ToListAsync();

        //    return _context.PreProcGeneralInfo2s != null ?
        //View(await preproc.OrderByDescending(t => t.Id).ToListAsync()) :
        //Problem("Entity set 'TaxFormReaderContext.TrTaxes'  is null.");
        //}

        //public async Task<IActionResult> ViewDetail(int id)
        //{
        //    var preproc = from t in _context.PreProcGeneralInfo2s
        //                  select t;
        //    return _context.PreProcGeneralInfo2s != null ?
        //View(await preproc.Where(t=> t.Id == id).OrderByDescending(t => t.Id).ToListAsync()) :
        //Problem("Entity set 'TaxFormReaderContext.TrTaxes'  is null.");
        //    //return View();
        //}
        
        public async Task<IActionResult> ViewDetail(int id)
        {
            getUserNow();
            if (id < 1 || _context.PreProcGeneralInfo2s == null)
            {
                return NotFound();
            }

            var preprocGI = await _context.PreProcGeneralInfo2s
                .FirstOrDefaultAsync(m => m.Id == id);
            if (preprocGI == null)
            {
                return NotFound();
            }

            var emp = await _context.VwEmployeeBaseEreqs.OrderBy(x => x.Nama).ToListAsync();

            var useremail = getUserNowEmail();

            var paramemail = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Current Login Email").FirstOrDefaultAsync();

            if (paramemail.AlphaNumericValue != null && paramemail.AlphaNumericValue != "")
            {
                useremail = paramemail.AlphaNumericValue;
            }

            int found = 0;

            if (found != 1)
            {
                ViewBag.useraccessSuperAdmin = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title.Equals("PreProc") && x.Parameter.Equals("Super Admin Access") && x.AlphaNumericValue.ToLower().Equals(useremail.ToLower()));
                if (ViewBag.useraccessSuperAdmin != null)
                {
                    found = 1;
                }
            }

            if (found != 1)
            {
                ViewBag.useraccesspm = await _context.VwGetPmmemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
                if (ViewBag.useraccesspm != null)
                {
                    found = 1;
                }
            }
            if (found != 1)
            {
                ViewBag.useraccesspc = await _context.VwGetPcmemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
                if (ViewBag.useraccesspc != null)
                {
                    found = 1;
                }
            }
            if (found != 1)
            {
                ViewBag.useraccesspmo = await _context.VwGetPmomemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
                if (ViewBag.useraccessproc != null)
                {
                    found = 1;
                }
                else
                {
                    ViewBag.useraccesspmohead = await _context.VwGetPmomemberFulls.FirstOrDefaultAsync(x => x.JobTitleLevel.Contains("Head") && x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
                }
            }

            ViewBag.attachRFP = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 2).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

            ViewBag.attachBOQ = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 6).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

            ViewBag.attachQuot = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 4).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

            ViewBag.attachPropo = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 5).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

            ViewBag.attachGP = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 15).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

            ViewBag.attachOthers = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 10).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();
            ViewBag.attachSPK = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 3).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();
            ViewBag.attachNego = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 7).OrderByDescending(t => t.Id)
                .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

            PreProcContextProcedures dbsp = new PreProcContextProcedures(_context);

            //ViewBag.principles = _context.VwVendors.OrderBy(x => x.Vendor).ToList();
            ViewBag.itemdetail = await _context.PreProcDetails.Where(x => x.Id == id && x.Idnew != "EDIT").OrderBy(x => x.DetailId).ToListAsync();
            ViewBag.fileattach = _context.PreProcFileAttachments.Where(x => x.PreProcId == id && x.HasBeenUploaded != false).OrderByDescending(x => x.Id).ToList();
            var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Detail");
            var commenttype2 = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Risk");
            ViewBag.approvalhistrisk = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == id && !x.Status.Contains("temp") && x.Type == commenttype2.Id).OrderByDescending(x => x.Id).ToListAsync();
            ViewBag.approvalhist = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == id && !x.Status.Contains("temp") && x.Type == commenttype.Id).OrderByDescending(x => x.Id).ToListAsync();
            ViewBag.pmocheck = await _context.PreProcPmochecklists.FirstOrDefaultAsync(x => x.PreProcId == id);
            //ViewBag.risktype = _context.MsRiskTypes.OrderByDescending(x => x.Id).ToList();

            //ViewBag.contingency = _context.MsContingencyCosts.OrderBy(x => x.Id).ToList();
            //ViewBag.principalpreload = _context.MsPrincipalPreloads.OrderBy(x => x.Id).ToList();


            ViewBag.riskdata = await _context.PreProcRiskData.Where(x => x.PreProcId == id && !x.CheckedBy.Contains("UncheckedBy")).Join(_context.PreProcRisks,
                        p => p.RiskId, d => d.RiskId,
                        (riskdata, risk) => new
                        {
                            PreProcId = riskdata.PreProcId
                          ,
                            RiskOrder = riskdata.RiskOrder
                          ,
                            CheckedBy = riskdata.CheckedBy
                          ,
                            RiskType = riskdata.RiskType
                          ,
                            RiskId = riskdata.RiskId
                            ,
                            RiskName = risk.RiskName
                          ,
                            RiskWeight = risk.RiskWeight
                          ,
                            IsActive = risk.IsActive
                        }).Where(x => x.IsActive == true).ToListAsync();
            var riskassess = await _context.PreProcRiskAssesmentData.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id);
            ViewBag.riskassess = riskassess;
            if (riskassess != null)
            {
                ViewBag.sourcebudget = _context.MsSourceBudgets.OrderBy(x => x.Id).FirstOrDefault(x=>x.Id == Int32.Parse(riskassess.SourceBudget));
                ViewBag.competition = _context.MsCompetitions.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskassess.Competition));
                ViewBag.proctype = _context.MsProcurementTypes.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskassess.ProcurementType));
                ViewBag.projectsource = _context.MsProjectSources.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskassess.ProjectSource));

            }

            var temprisk = await _context.PreProcRiskAssesmentData.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id);
            if (temprisk != null)
            {
                if (temprisk.EmailAssignTo != null && temprisk.EmailAssignTo != "")
                {
                    var splitemail = temprisk.EmailAssignTo.Split(";");
                    ViewBag.riskassignto = "";
                    for (int i = 0; i < splitemail.Length; i++)
                    {
                        if (splitemail[i] != "")
                        {
                            var tes = await _context.VwEmployeeBaseEreqs.FirstOrDefaultAsync(x => x.EmailAddress.Equals(splitemail[i]));
                            ViewBag.riskassignto += tes.Nama + "; ";
                        }
                    }
                }

            }

            ViewBag.projectrevmlpt = await _context.PreProcPrincipalProjectRevenues.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id && x.PrincipalName.Equals("MLPT"));
            ViewBag.projectrev = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == id && x.PrincipalName != "MLPT").OrderByDescending(x => x.Id).ToListAsync();
            ViewBag.countsubcon = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == id && x.PrincipalName != "MLPT").OrderByDescending(x => x.Id).CountAsync();
            ViewBag.datapo = await _context.PreProcPodata.Where(x => x.PreProcId == preprocGI.Id).ToListAsync();
            ViewBag.customers = _context.MsCustomerVws.OrderBy(x => x.AccountName).FirstOrDefault(x=>x.InitialCode == preprocGI.Customer);
            var billreason = await _context.PreProcPendingBillingReasons.FirstOrDefaultAsync(x => x.PreProcId == id && x.NeedEscalation == true);
            ViewBag.billreason = billreason;
            if (billreason != null)
            {
                ViewBag.escalatetoam = await _context.MsEscalationToAms.OrderBy(x => x.Id).FirstOrDefaultAsync(x => x.Id == billreason.EscalationReasonId);
            }

            if (preprocGI.EmailCommitteeDefault != "" && preprocGI.EmailCommitteeDefault != null)
            {
                var tempbod = preprocGI.EmailCommitteeDefault.Split(";");
                var bod = "";
                for (int i = 0; i < tempbod.Length; i++)
                {
                    if (tempbod[i] != "")
                    {
                        var tempbod2 = await _context.VwGetBodfulls.FirstOrDefaultAsync(x => x.EmailAddress == tempbod[i]);
                        bod += tempbod2.Nama + ";";
                    }

                }
                ViewBag.bodname = bod;
            }
            var totpriceidr = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("IDR")).Select(x => decimal.Parse(x.TotalPrice)).ToListAsync();
            var temptotpriceidr = totpriceidr.Sum().ToString("#,###");

            var totpriceusd = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("USD")).Select(x => decimal.Parse(x.TotalPrice)).ToListAsync();
            var temptotpriceusd = totpriceusd.Sum().ToString("#,###");

            var totcogsidr = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("IDR")).Select(x => decimal.Parse(x.TotalCogs)).ToListAsync();
            var temptotcogsidr = totcogsidr.Sum().ToString("#,###");

            var totcogsusd = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("USD")).Select(x => decimal.Parse(x.TotalCogs)).ToListAsync();
            var temptotcogsusd = totcogsusd.Sum().ToString("#,###");

            ViewBag.totalcogsidr = temptotcogsidr + ",00";
            ViewBag.totalcogsusd = temptotcogsusd + ",00";
            ViewBag.totpriceidr = temptotpriceidr + ",00";
            ViewBag.totpriceusd = temptotpriceusd + ",00";

            return View(preprocGI);
            //return View(Tuple.Create(preprocGI, emp));
        }
        
        public FileResult GetFileAttachment(string target, string downVerRFP, string downVerBOQ, string downVerQuot, string downVerGP, string downVerPropo, string downVerOthers)
        {
            getUserNow();
            //string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files");

            //string fileNameWithPath = Path.Combine(path, target);//file.FileName);

            if (target != null)
            {
                var tempTarget = target.Split(";");
                byte[] FileBytes = System.IO.File.ReadAllBytes(tempTarget[0]);

                //"application/force-download" opsi lain
                return File(FileBytes, "application/octet-stream", tempTarget[1]);
            }
            else
            {
                var tempfile = "";

                if (downVerRFP != null)
                {
                    tempfile = downVerRFP;
                }
                else if (downVerBOQ != null)
                {
                    tempfile = downVerBOQ;
                }
                else if (downVerQuot != null)
                {
                    tempfile = downVerQuot;
                }
                else if (downVerGP != null)
                {
                    tempfile = downVerGP;
                }
                else if (downVerPropo != null)
                {
                    tempfile = downVerPropo;
                }
                else if (downVerOthers != null)
                {
                    tempfile = downVerOthers;
                }

                var tempTarget = tempfile.Split("\\");
                byte[] FileBytes = System.IO.File.ReadAllBytes(tempfile);

                //"application/force-download" opsi lain
                return File(FileBytes, "application/octet-stream", tempTarget.Last());
            }
        }

        public async Task<IActionResult> EditDetail(int id)
        {
            getUserNow();

            var useremail = getUserNowEmail().ToLower();
            var paramemail = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Current Login Email").FirstOrDefaultAsync();

            if (paramemail.AlphaNumericValue != null && paramemail.AlphaNumericValue != "")
            {
                useremail = paramemail.AlphaNumericValue;
            }

            //useremail = "ferianto.hanafie@multipolar.com";    //PROC MEMBER
            int found = 0;

            // Fetch necessary user data once
            var vwGetAmmembers = await _context.VwGetAmmembers.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetBodfulls = await _context.VwGetBodfulls.Where(x => x.EmailAddress.ToLower() == useremail).ToListAsync();
            var vwGetTenderAdminMemberFulls = await _context.VwGetTenderAdminMemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetPresalesMemberFulls = await _context.VwGetPresalesMemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetPossibleProductManagers = await _context.VwGetPossibleProductManagers.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetPmmemberFulls = await _context.VwGetPmmemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetPcmemberFulls = await _context.VwGetPcmemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetProjectSupportMemberFulls = await _context.VwGetProjectSupportMemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetPmomemberFulls = await _context.VwGetPmomemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetProcurementMemberFulls = await _context.VwGetProcurementMemberFulls.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var vwGetWarehouseMembers = await _context.VwGetWarehouseMembers.Where(x => x.EmployeeEmail.ToLower() == useremail).ToListAsync();
            var msParameterValues = await _context.MsParameterValues.Where(x => x.AlphaNumericValue.ToLower() == useremail && x.Title == "PreProc" && x.Parameter == "Super Admin Access").ToListAsync();

            // Set ViewBags
            ViewBag.useraccessam = vwGetAmmembers.FirstOrDefault();
            ViewBag.useraccessbod = vwGetBodfulls.FirstOrDefault();
            

            // Determine access level
            if (vwGetAmmembers.Any())
            {
                ViewBag.userammanager = vwGetAmmembers.FirstOrDefault(x => x.JobTitle.Contains("Head"));
                found = 1;
            }
            else if (msParameterValues.Any())
            {
                ViewBag.useraccessSuperAdmin = msParameterValues.FirstOrDefault();
                found = 1;
            }
            else if (vwGetTenderAdminMemberFulls.Any())
            {
                ViewBag.useraccessta = vwGetTenderAdminMemberFulls.FirstOrDefault();
                found = 1;
            }
            else if (vwGetPresalesMemberFulls.Any())
            {
                ViewBag.useraccesspresales = vwGetPresalesMemberFulls.FirstOrDefault();
                found = 1;
            }
            else if (vwGetPossibleProductManagers.Any())
            {
                ViewBag.useraccessprodmgr = vwGetPossibleProductManagers.FirstOrDefault();
                found = 1;
            }
            else if (vwGetPmmemberFulls.Any())
            {
                ViewBag.useraccesspm = vwGetPmmemberFulls.FirstOrDefault();
                ViewBag.useraccesspmmgr = vwGetPmmemberFulls.FirstOrDefault(x => x.JobTitleLevel.Contains("Department Head"));
                found = 1;
            }
            else if (vwGetPcmemberFulls.Any())
            {
                ViewBag.useraccesspc = vwGetPcmemberFulls.FirstOrDefault();
                ViewBag.useraccesspcmgr = vwGetPcmemberFulls.FirstOrDefault(x => x.JobTitleLevel.Contains("Section Head"));
                found = 1;
            }
            else if (vwGetProjectSupportMemberFulls.Any())
            {
                ViewBag.useraccessps = vwGetProjectSupportMemberFulls.FirstOrDefault();
                ViewBag.useraccesspsmgr = vwGetProjectSupportMemberFulls.FirstOrDefault(x => x.JobTitleLevel.Contains("Section Head"));
                found = 1;
            }
            else if (vwGetPmomemberFulls.Any())
            {
                ViewBag.useraccesspmo = vwGetPmomemberFulls.FirstOrDefault();
                ViewBag.useraccesspmohead = vwGetPmomemberFulls.FirstOrDefault(x => x.JobTitleLevel.Contains("Head"));
                found = 1;
            }
            else if (vwGetProcurementMemberFulls.Any())
            {
                ViewBag.useraccessproc = vwGetProcurementMemberFulls.FirstOrDefault();
                found = 1;
            }
            else if (vwGetWarehouseMembers.Any())
            {
                ViewBag.useraccesswh = vwGetWarehouseMembers.FirstOrDefault();
            }

            if (id < 1 || _context.PreProcGeneralInfo2s == null)
            {
                return NotFound();
            }

            var preprocGI = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(m => m.Id == id);
            if (preprocGI == null)
            {
                return NotFound();
            }

            if (preprocGI.PicProc != null && preprocGI.Stage == "Delivery")
            {
                var tempsplitpic = preprocGI.PicProc.Split(";");

                var temppicname = "";

                foreach (var item in tempsplitpic)
                {
                    if (item != null && item != "" )
                    {
                        temppicname += _context.VwEmployeeBases.FirstOrDefault(x => x.EmployeeEmail == item).EmployeeName + "; ";
                    }
                    
                }
                ViewBag.picname = temppicname;
            }


            var empbaseereq = await _context.VwEmployeeBaseEreqs.Where(x => x.EmailAddress != null).OrderBy(x => x.Nama).ToListAsync();


            if (preprocGI.Stage == "RiskAssessment")
            {
                var riskassessdata = await _context.PreProcRiskAssesmentData.FirstOrDefaultAsync(x => x.PreProcId == id);
                if (riskassessdata != null)
                {
                    var split = riskassessdata.EmailAssignTo.Split(";");
                    ViewBag.useraccesscommittee = split.FirstOrDefault(item => item.ToLower() == useremail);
                }
                ViewBag.riskdata = await _context.PreProcRiskData.Where(x => x.PreProcId == id && !x.CheckedBy.Contains("UncheckedBy")).Join(_context.PreProcRisks,
                        p => p.RiskId, d => d.RiskId,
                        (riskdata, risk) => new
                        {
                            PreProcId = riskdata.PreProcId
                          ,
                            RiskOrder = riskdata.RiskOrder
                          ,
                            CheckedBy = riskdata.CheckedBy
                          ,
                            RiskType = riskdata.RiskType
                          ,
                            RiskId = riskdata.RiskId
                            ,
                            RiskName = risk.RiskName
                          ,
                            RiskWeight = risk.RiskWeight
                          ,
                            IsActive = risk.IsActive
                        }).Where(x => x.IsActive == true).ToListAsync();
                ViewBag.riskassess = await _context.PreProcRiskAssesmentData.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id);
                var temprisk = await _context.PreProcRiskAssesmentData.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id);
                if (temprisk != null)
                {
                    if (temprisk.EmailAssignTo != null && temprisk.EmailAssignTo != "")
                    {
                        var splitemail = temprisk.EmailAssignTo.Split(";");
                        ViewBag.riskassignto = "";
                        for (int i = 0; i < splitemail.Length; i++)
                        {
                            if (splitemail[i] != "")
                            {
                                var tes = await _context.VwEmployeeBaseEreqs.FirstOrDefaultAsync(x => x.EmailAddress.Equals(splitemail[i]));
                                ViewBag.riskassignto += tes.Nama + "; ";
                            }
                        }
                    }

                }

                ViewBag.risk = _context.PreProcRisks.ToList();
                ViewBag.projectrevmlpt = await _context.PreProcPrincipalProjectRevenues.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id && x.PrincipalName.Equals("MLPT"));
                ViewBag.projectrev = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == id && x.PrincipalName != "MLPT").OrderByDescending(x => x.Id).ToListAsync();
                ViewBag.countsubcon = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == id && x.PrincipalName != "MLPT").OrderByDescending(x => x.Id).CountAsync();
                ViewBag.risktype = _context.MsRiskTypes.OrderByDescending(x => x.Id).ToList();
                ViewBag.sourcebudget = _context.MsSourceBudgets.OrderBy(x => x.Id).ToList();
                ViewBag.competition = _context.MsCompetitions.OrderBy(x => x.Id).ToList();
                ViewBag.contingency = _context.MsContingencyCosts.OrderBy(x => x.Id).ToList();
                ViewBag.proctype = _context.MsProcurementTypes.OrderBy(x => x.Id).ToList();
                ViewBag.projectsource = _context.MsProjectSources.OrderBy(x => x.Id).ToList();
                ViewBag.principalpreload = _context.MsPrincipalPreloads.OrderBy(x => x.Id).ToList();
                ViewBag.empbaseereq = empbaseereq.Where(x => x.Status.Contains("m")).ToList();
            }
            else if (preprocGI.Stage == "Delivery")
            {
                ViewBag.billreason = await _context.PreProcPendingBillingReasons.FirstOrDefaultAsync(x => x.PreProcId == id && x.NeedEscalation);
            }
            ViewBag.picreason = empbaseereq;

            var attachTypes = new List<int> { 2, 6, 4, 5, 15, 10, 3, 7 };
            var attachments = await _context.PreProcHeaderAttachments
                .Where(t => t.PreProcId == preprocGI.Id && attachTypes.Contains(t.DocumentTypeId.Value))
                .OrderByDescending(t => t.Id)
                .ToListAsync();

            ViewBag.attachRFP = attachments.Where(t => t.DocumentTypeId == 2).ToList();
            ViewBag.attachBOQ = attachments.Where(t => t.DocumentTypeId == 6).ToList();
            ViewBag.attachQuot = attachments.Where(t => t.DocumentTypeId == 4).ToList();
            ViewBag.attachPropo = attachments.Where(t => t.DocumentTypeId == 5).ToList();
            ViewBag.attachGP = attachments.Where(t => t.DocumentTypeId == 15).ToList();
            ViewBag.attachOthers = attachments.Where(t => t.DocumentTypeId == 10).ToList();
            ViewBag.attachSPK = attachments.Where(t => t.DocumentTypeId == 3).ToList();
            ViewBag.attachNego = attachments.Where(t => t.DocumentTypeId == 7).ToList();

            // Set additional ViewBags
            ViewBag.amemployees = await _context.VwGetAmmembers.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.technicalemployees = await _context.VwGetPresalesMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.tenderemployees = await _context.VwGetTenderAdminMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.customers = _context.MsCustomerVws.OrderBy(x => x.AccountName).ToList();
            ViewBag.picprocs = _context.VwGetProcurementMemberFulls.OrderBy(x => x.EmployeeName).ToList();
            ViewBag.principles = _context.VwVendors.OrderBy(x => x.Vendor).ToList();
            ViewBag.itemdescs = _context.VwGetNkpnames.Where(x => x.VendorId.Equals("Acer")).OrderBy(x => x.Nkpname).ToList();
            ViewBag.itemdescsall = _context.VwGetNkpnames.OrderBy(x => x.Nkpname).ToList();
            ViewBag.pm = await _context.VwGetPmmemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.pc = await _context.VwGetPcmemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.escalatetoam = await _context.MsEscalationToAms.OrderBy(x => x.Id).ToListAsync();
            ViewBag.empbaseereq = await _context.VwEmployeeBaseEreqs.Where(x => x.Status.Contains("m")).OrderBy(x => x.Nama).ToListAsync();
            ViewBag.bod = await _context.VwGetBodfulls.OrderBy(x => x.Nama).ToListAsync();
            ViewBag.warehouse = await _context.VwGetWarehouseMembers.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.finance = await _context.VwEmployeeBaseEreqs.Where(x => x.CurrentGroup == "CSAF").OrderBy(x => x.Nama).ToListAsync();
            ViewBag.financehead = await _context.VwEmployeeBaseEreqs.Where(x => x.CurrentGroup == "CSAF" && x.Status == "m").OrderBy(x => x.Nama).ToListAsync();
            ViewBag.projectsupp = await _context.VwGetProjectSupportMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
            ViewBag.projecttypes = _context.MsProjectTypes.ToList();
            ViewBag.presales = await _context.VwGetAllOptyIds.OrderByDescending(x => x.Id).ToListAsync();
            ViewBag.pids = await _context.VwGetAllPids.OrderByDescending(x => x.TrProjectIdId).ToListAsync();
            ViewBag.items = _context.MsPreProcStages.ToList();
            ViewBag.itemdetail = await _context.PreProcDetails.Where(x => x.Id == id && x.Idnew != "EDIT").OrderBy(x => x.DetailId).ToListAsync();
            ViewBag.fileattach = _context.PreProcFileAttachments.Where(x => x.PreProcId == id && x.HasBeenUploaded != false).OrderByDescending(x => x.Id).ToList();
            ViewBag.pmocheck = await _context.PreProcPmochecklists.FirstOrDefaultAsync(x => x.PreProcId == id);
            var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Detail");
            var commenttype2 = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Risk");
            ViewBag.approvalhistrisk = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == id && !x.Status.Contains("temp") && x.Type == commenttype2.Id).OrderByDescending(x => x.Id).ToListAsync();
            ViewBag.approvalhist = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == id && !x.Status.Contains("temp") && x.Type == commenttype.Id).OrderByDescending(x => x.Id).ToListAsync();
            ViewBag.datapo = await _context.PreProcPodata.Where(x => x.PreProcId == preprocGI.Id).ToListAsync();
            ViewBag.prodmgr = _context.VwGetNkpnames.Select(x => new
            {
                x.ProductManager,
                x.Domain
            }).Distinct().ToList();
            ViewBag.urlgenpid = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title.Equals("PreProc") && x.Parameter.Equals("Url Generate PID"));

            var totpriceidr = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("IDR") && x.TotalPrice != null).Select(x => decimal.Parse(x.TotalPrice)).ToListAsync();
            var temptotpriceidr = totpriceidr.Sum().ToString("#,###");

            var totpriceusd = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("USD") && x.TotalPrice != null).Select(x => decimal.Parse(x.TotalPrice)).ToListAsync();
            var temptotpriceusd = totpriceusd.Sum().ToString("#,###");

            var totcogsidr = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("IDR") && x.TotalCogs != null).Select(x => decimal.Parse(x.TotalCogs)).ToListAsync();
            var temptotcogsidr = totcogsidr.Sum().ToString("#,###");

            var totcogsusd = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("USD") && x.TotalCogs != null).Select(x => decimal.Parse(x.TotalCogs)).ToListAsync();
            var temptotcogsusd = totcogsusd.Sum().ToString("#,###");

            ViewBag.totalcogsidr = temptotcogsidr + ",00";
            ViewBag.totalcogsusd = temptotcogsusd + ",00";
            ViewBag.totpriceidr = temptotpriceidr + ",00";
            ViewBag.totpriceusd = temptotpriceusd + ",00";

            var picname = "";

            if (preprocGI.PicProc != "" && preprocGI.PicProc != null)
            {
                var tempsplit = preprocGI.PicProc.Split(";");

                foreach (var item in tempsplit)
                {
                    if (item != "")
                    {
                        var tempnamepic = await _context.VwGetProcurementMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower() == item.ToLower());
                        if (tempnamepic != null)
                        {
                            picname += tempnamepic.EmployeeName + ";";
                        }
                    }

                }
            }

            ViewBag.existpicproc = picname;

            if (preprocGI.EmailCommitteeDefault != "" && preprocGI.EmailCommitteeDefault != null)
            {
                var tempbod = preprocGI.EmailCommitteeDefault.Split(";");
                var bod = "";
                for (int i = 0; i < tempbod.Length; i++)
                {
                    if (tempbod[i] != "")
                    {
                        var tempbod2 = await _context.VwGetBodfulls.FirstOrDefaultAsync(x => x.EmailAddress == tempbod[i]);
                        bod += tempbod2.Nama + ";";
                    }

                }
                ViewBag.bodname = bod;
            }

            return View(preprocGI);
        }


        //public async Task<IActionResult> EditDetail(int id)
        //{
        //    getUserNow();

        //    var useremail = getUserNowEmail();
        //    //useremail = "hananto.wibowo@multipolar.com";      //PM MEMBER
        //    //useremail = "kevin.christian@multipolar.com";     //AM MEMBER
        //    //useremail = "ferianto.hanafie@multipolar.com";    //PROC MEMBER
        //    //useremail = "dini.hayati@multipolar.com";           //PS MEMBER
        //    //useremail = "solehudin@multipolar.com";           //WH MEMBER
        //    //useremail = "arief@multipolar.com";           //TEST COMMITTEE MEMBER
        //    //useremail = "IVAN@multipolar.com";           //BOD MEMBER

        //    int found = 0;

        //    ViewBag.useraccessam = await _context.VwGetAmmembers.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //    ViewBag.useraccessbod = await _context.VwGetBodfulls.FirstOrDefaultAsync(x => x.EmailAddress.ToLower().Equals(useremail.ToLower()));
        //    var useraccessam = await _context.VwGetAmmembers.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));

        //    if (useraccessam != null)
        //    {
        //        ViewBag.userammanager = await _context.VwGetAmmembers.FirstOrDefaultAsync(x => x.JobTitle.Contains("Head") && x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        found = 1;
        //    }

        //    if (found != 1)
        //    {
        //        ViewBag.useraccessSuperAdmin = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title.Equals("PreProc") && x.Parameter.Equals("Super Admin Access") && x.AlphaNumericValue.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccessSuperAdmin != null)
        //        {
        //            found = 1;
        //        }
        //        else
        //        {
        //            ViewBag.useraccessta = await _context.VwGetTenderAdminMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //            if (ViewBag.useraccessta != null)
        //            {
        //                found = 1;
        //            }
        //        }
        //    }

        //    if (found != 1)
        //    {
        //        ViewBag.useraccesspresales = await _context.VwGetPresalesMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccesspresales != null)
        //        {
        //            found = 1;
        //        }
        //        else
        //        {
        //            ViewBag.useraccessprodmgr = await _context.VwGetPossibleProductManagers.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //            if (ViewBag.useraccessprodmgr != null)
        //            {
        //                found = 1;
        //            }
        //        }
        //    }
        //    if (found != 1)
        //    {
        //        ViewBag.useraccesspm = await _context.VwGetPmmemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccesspm != null)
        //        {
        //            found = 1;
        //        }
        //        else
        //        {
        //            ViewBag.useraccesspmmgr = await _context.VwGetPmmemberFulls.FirstOrDefaultAsync(x => x.JobTitleLevel == "Department Head" && x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //            if (ViewBag.useraccesspmmgr != null)
        //            {
        //                found = 1;
        //            }
        //        }
        //    }
        //    if (found != 1)
        //    {
        //        ViewBag.useraccesspc = await _context.VwGetPcmemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccesspc != null)
        //        {
        //            found = 1;
        //        }
        //        else
        //        {
        //            ViewBag.useraccesspcmgr = await _context.VwGetPcmemberFulls.FirstOrDefaultAsync(x => x.JobTitleLevel == "Section Head" && x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        }

        //    }
        //    if (found != 1)
        //    {
        //        ViewBag.useraccessps = await _context.VwGetProjectSupportMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccessps != null)
        //        {
        //            found = 1;
        //        }
        //        else
        //        {
        //            ViewBag.useraccesspsmgr = await _context.VwGetProjectSupportMemberFulls.FirstOrDefaultAsync(x => x.JobTitleLevel == "Section Head" && x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        }

        //    }
        //    if (found != 1)
        //    {
        //        ViewBag.useraccesspmo = await _context.VwGetPmomemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccessproc != null)
        //        {
        //            found = 1;
        //        }
        //        else
        //        {
        //            ViewBag.useraccesspmohead = await _context.VwGetPmomemberFulls.FirstOrDefaultAsync(x => x.JobTitleLevel.Contains("Head") && x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        }
        //    }
        //    if (found != 1)
        //    {
        //        ViewBag.useraccessproc = await _context.VwGetProcurementMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        if (ViewBag.useraccessproc == null)
        //        {
        //            ViewBag.useraccesswh = await _context.VwGetWarehouseMembers.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
        //        }
        //    }

        //    if (id < 1 || _context.PreProcGeneralInfo2s == null)
        //    {
        //        return NotFound();
        //    }

        //    var preprocGI = await _context.PreProcGeneralInfo2s
        //        .FirstOrDefaultAsync(m => m.Id == id);
        //    if (preprocGI == null)
        //    {
        //        return NotFound();
        //    }

        //    if (preprocGI.Stage == "RiskAssessment")
        //    {
        //        var riskassessdata = await _context.PreProcRiskAssesmentData.FirstOrDefaultAsync(x=> x.PreProcId == id);

        //        if (riskassessdata != null)
        //        {
        //            var split = riskassessdata.EmailAssignTo.Split(";");

        //            foreach (var item in split)
        //            {
        //                if (item.ToLower() == useremail.ToLower())
        //                {
        //                    ViewBag.useraccesscommittee = item;
        //                    break;
        //                }
        //            }


        //        }
        //    }

        //    ViewBag.attachRFP = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 2).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

        //    ViewBag.attachBOQ = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 6).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

        //    ViewBag.attachQuot = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 4).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

        //    ViewBag.attachPropo = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 5).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

        //    ViewBag.attachGP = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 15).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

        //    ViewBag.attachOthers = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 10).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();
        //    ViewBag.attachSPK = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 3).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();
        //    ViewBag.attachNego = await _context.PreProcHeaderAttachments.Where(t => t.DocumentTypeId == 7).OrderByDescending(t => t.Id)
        //        .Where(m => m.PreProcId == preprocGI.Id).ToListAsync();

        //    PreProcContextProcedures dbsp = new PreProcContextProcedures(_context);

        //    //var preprocdetail = await _context.PreProcDetails.Where(x => x.Id == id).OrderBy(x => x.DetailId).ToListAsync();

        //    //var preprocdetaildel = await _context.TempPreProcDetails.Where(x => x.Id == id && x.DetailIdold != null).OrderBy(x => x.DetailId).ToListAsync();

        //    //_context.RemoveRange(preprocdetaildel);
        //    //await _context.SaveChangesAsync();

        //    //TempPreProcDetail tempdetail;

        //    //foreach (var item in preprocdetail)
        //    //{
        //    //    tempdetail = new TempPreProcDetail();

        //    //    tempdetail.Id = item.Id;
        //    //    tempdetail.DetailIdold = item.DetailId;
        //    //    tempdetail.Principle = item.Principle;
        //    //    tempdetail.ItemDesc = item.ItemDesc;
        //    //    tempdetail.Spec = item.Spec;
        //    //    tempdetail.SurDukCheck = item.SurDukCheck;
        //    //    tempdetail.SurDuk = item.SurDuk;
        //    //    tempdetail.Qty = item.Qty;
        //    //    tempdetail.Currency = item.Currency;
        //    //    tempdetail.UnitCogs = item.UnitCogs;
        //    //    tempdetail.UnitPrice = item.UnitPrice;
        //    //    tempdetail.TotalCogs = item.TotalCogs;
        //    //    tempdetail.TotalPrice = item.TotalPrice;
        //    //    tempdetail.Vendor = item.Vendor;
        //    //    tempdetail.VendorQuotation = item.VendorQuotation;
        //    //    tempdetail.ExpectedPoissuedDate = item.ExpectedPoissuedDate;
        //    //    tempdetail.Mlptpono = item.Mlptpono;
        //    //    tempdetail.Mlptpodate = item.Mlptpodate;
        //    //    tempdetail.Mlptpo = item.Mlptpo;
        //    //    tempdetail.Eta = item.Eta;
        //    //    tempdetail.Grdate = item.Grdate;
        //    //    tempdetail.Grqty = item.Grqty;
        //    //    tempdetail.Grbacklog = item.Grbacklog;
        //    //    tempdetail.PresalesReview = item.PresalesReview;
        //    //    tempdetail.AmmanagerApproval = item.AmmanagerApproval;
        //    //    tempdetail.RemarkOnItem = item.RemarkOnItem;
        //    //    tempdetail.ProcessedPresalesReview = item.ProcessedPresalesReview;
        //    //    tempdetail.ProcessedAmreview = item.ProcessedAmreview;
        //    //    tempdetail.PonumberId = item.PonumberId;
        //    //    tempdetail.Nkpcode = item.Nkpcode;
        //    //    tempdetail.ProductManager = item.ProductManager;
        //    //    tempdetail.Warranty = item.Warranty;
        //    //    tempdetail.Insurance = item.Insurance;
        //    //    tempdetail.Risk = item.Risk;
        //    //    tempdetail.AuditDetailId = item.AuditDetailId;
        //    //    tempdetail.Idnew = item.Idnew;
        //    //    tempdetail.Grn = item.Grn;
        //    //    tempdetail.Pidfull = item.Pidfull;
        //    //    tempdetail.FirstId = item.FirstId;
        //    //    tempdetail.SecondId = item.SecondId;

        //    //    _context.Add(tempdetail);
        //    //    await _context.SaveChangesAsync();
        //    //}

        //    ViewBag.amemployees = await _context.VwGetAmmembers.OrderBy(x => x.EmployeeName).ToListAsync();
        //    ViewBag.technicalemployees = await _context.VwGetPresalesMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
        //    ViewBag.tenderemployees = await _context.VwGetTenderAdminMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
        //    ViewBag.customers = _context.MsCustomerVws.OrderBy(x => x.AccountName).ToList();
        //    ViewBag.picprocs = _context.VwGetProcurementMemberFulls.OrderBy(x => x.EmployeeName).ToList();
        //    ViewBag.principles = _context.VwVendors.OrderBy(x => x.Vendor).ToList();
        //    ViewBag.itemdescs = _context.VwGetNkpnames.Where(x => x.VendorId.Equals("Acer")).OrderBy(x => x.Nkpname).ToList();
        //    ViewBag.itemdescsall = _context.VwGetNkpnames.OrderBy(x => x.Nkpname).ToList();
        //    ViewBag.projecttypes = _context.MsProjectTypes.ToList();
        //    ViewBag.presales = await _context.VwGetAllOptyIds.OrderByDescending(x=>x.Id).ToListAsync();
        //    ViewBag.pids = await _context.VwGetAllPids.OrderByDescending(x => x.TrProjectIdId).ToListAsync();
        //    ViewBag.items = _context.MsPreProcStages.ToList();
        //    ViewBag.itemdetail = await _context.PreProcDetails.Where(x=>x.Id == id && x.Idnew != "EDIT").OrderBy(x=>x.DetailId).ToListAsync();
        //    ViewBag.fileattach = _context.PreProcFileAttachments.Where(x => x.PreProcId == id && x.HasBeenUploaded != false).OrderByDescending(x => x.Id).ToList();
        //    var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Detail");
        //    var commenttype2 = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Risk");
        //    ViewBag.approvalhistrisk = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == id && !x.Status.Contains("temp") && x.Type == commenttype2.Id).OrderByDescending(x => x.Id).ToListAsync();
        //    ViewBag.approvalhist = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == id && !x.Status.Contains("temp") && x.Type == commenttype.Id).OrderByDescending(x => x.Id).ToListAsync();
        //    ViewBag.pmocheck = await _context.PreProcPmochecklists.FirstOrDefaultAsync(x => x.PreProcId == id);
        //    ViewBag.risktype = _context.MsRiskTypes.OrderByDescending(x=>x.Id).ToList();
        //    ViewBag.sourcebudget = _context.MsSourceBudgets.OrderBy(x => x.Id).ToList();
        //    ViewBag.competition = _context.MsCompetitions.OrderBy(x => x.Id).ToList();
        //    ViewBag.contingency = _context.MsContingencyCosts.OrderBy(x => x.Id).ToList();
        //    ViewBag.proctype = _context.MsProcurementTypes.OrderBy(x => x.Id).ToList();
        //    ViewBag.projectsource = _context.MsProjectSources.OrderBy(x => x.Id).ToList();
        //    ViewBag.principalpreload = _context.MsPrincipalPreloads.OrderBy(x => x.Id).ToList();
        //    ViewBag.empbaseereq = await _context.VwEmployeeBaseEreqs.Where(x => x.Status.Contains("m")).OrderBy(x => x.Nama).ToListAsync();
        //    ViewBag.prodmgr = _context.VwGetNkpnames.Select(x => new
        //    {
        //        x.ProductManager,
        //        x.Domain
        //    }).Distinct().ToList();

        //    ViewBag.riskdata = await _context.PreProcRiskData.Where(x => x.PreProcId == id && !x.CheckedBy.Contains("UncheckedBy")).Join(_context.PreProcRisks,
        //                p => p.RiskId, d => d.RiskId,
        //                (riskdata, risk) => new
        //                {
        //                    PreProcId = riskdata.PreProcId
        //                  ,
        //                    RiskOrder = riskdata.RiskOrder
        //                  ,
        //                    CheckedBy = riskdata.CheckedBy
        //                  ,
        //                    RiskType = riskdata.RiskType
        //                  ,
        //                    RiskId = riskdata.RiskId
        //                    ,
        //                    RiskName = risk.RiskName
        //                  ,
        //                    RiskWeight = risk.RiskWeight
        //                  ,
        //                    IsActive = risk.IsActive
        //                }).Where(x=>x.IsActive == true).ToListAsync();
        //    ViewBag.riskassess = await _context.PreProcRiskAssesmentData.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id);
        //    var temprisk = await _context.PreProcRiskAssesmentData.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id);
        //    if (temprisk != null)
        //    {
        //        if (temprisk.EmailAssignTo != null && temprisk.EmailAssignTo != "")
        //        {
        //            var splitemail = temprisk.EmailAssignTo.Split(";");
        //            ViewBag.riskassignto = "";
        //            for (int i = 0; i < splitemail.Length; i++)
        //            {
        //                if (splitemail[i] != "")
        //                {
        //                    var tes = await _context.VwEmployeeBaseEreqs.FirstOrDefaultAsync(x => x.EmailAddress.Equals(splitemail[i]));
        //                    ViewBag.riskassignto += tes.Nama + "; ";
        //                }
        //            }
        //        }

        //    }

        //    ViewBag.risk = _context.PreProcRisks.ToList();
        //    ViewBag.projectrevmlpt = await _context.PreProcPrincipalProjectRevenues.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == id && x.PrincipalName.Equals("MLPT"));
        //    ViewBag.projectrev = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == id && x.PrincipalName != "MLPT").OrderByDescending(x => x.Id).ToListAsync();
        //    ViewBag.countsubcon = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == id && x.PrincipalName != "MLPT").OrderByDescending(x => x.Id).CountAsync();
        //    ViewBag.projectsupp = await _context.VwGetProjectSupportMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
        //    ViewBag.pm = await _context.VwGetPmmemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
        //    ViewBag.pc = await _context.VwGetPcmemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
        //    //ViewBag.datacs = await _context.CogsrevNpcs.Where(x => x.ProjectId == preprocGI.Pid).ToListAsync();
        //    ViewBag.datapo = await _context.PreProcPodata.Where(x => x.PreProcId == preprocGI.Id).ToListAsync();
        //    ViewBag.escalatetoam = await _context.MsEscalationToAms.OrderBy(x=>x.Id).ToListAsync();
        //    ViewBag.warehouse = await _context.VwGetWarehouseMembers.OrderBy(x => x.EmployeeName).ToListAsync();
        //    ViewBag.finance = await _context.VwEmployeeBaseEreqs.Where(x=> x.CurrentGroup == "CSAF").OrderBy(x => x.Nama).ToListAsync();
        //    ViewBag.financehead = await _context.VwEmployeeBaseEreqs.Where(x => x.CurrentGroup == "CSAF" && x.Status == "m").OrderBy(x => x.Nama).ToListAsync();

        //    var picname = "";

        //    if (preprocGI.PicProc != "" && preprocGI.PicProc != null)
        //    {
        //        var tempsplit = preprocGI.PicProc.Split(";");

        //        foreach (var item in tempsplit)
        //        {
        //            if (item != "")
        //            {
        //                var tempnamepic = await _context.VwGetProcurementMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower() == item.ToLower());
        //                if (tempnamepic != null)
        //                {
        //                   picname += tempnamepic.EmployeeName + ";";
        //                }
        //            }

        //        }
        //    }

        //    ViewBag.existpicproc = picname;

        //    if (preprocGI.EmailCommitteeDefault != "" && preprocGI.EmailCommitteeDefault != null)
        //    {
        //        var tempbod = preprocGI.EmailCommitteeDefault.Split(";");
        //        var bod = "";
        //        for (int i = 0; i < tempbod.Length; i++)
        //        {
        //            if (tempbod[i] != "")
        //            {
        //                var tempbod2 = await _context.VwGetBodfulls.FirstOrDefaultAsync(x => x.EmailAddress == tempbod[i]);
        //                bod += tempbod2.Nama + ";";
        //            }

        //        }
        //        ViewBag.bodname = bod;
        //    }
        //    ViewBag.bod = await _context.VwGetBodfulls.OrderBy(x => x.Nama).ToListAsync();
        //    var totpriceidr = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("IDR") && x.TotalPrice != null).Select(x => decimal.Parse(x.TotalPrice)).ToListAsync();
        //    var temptotpriceidr = totpriceidr.Sum().ToString("#,###");

        //    var totpriceusd = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("USD") && x.TotalPrice != null).Select(x => decimal.Parse(x.TotalPrice)).ToListAsync();
        //    var temptotpriceusd = totpriceusd.Sum().ToString("#,###");

        //    var totcogsidr = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("IDR") && x.TotalCogs != null).Select(x => decimal.Parse(x.TotalCogs)).ToListAsync();
        //    var temptotcogsidr = totcogsidr.Sum().ToString("#,###");

        //    var totcogsusd = await _context.PreProcDetails.Where(x => x.Id == id && x.Currency.Equals("USD") && x.TotalCogs != null).Select(x => decimal.Parse(x.TotalCogs)).ToListAsync();
        //    var temptotcogsusd = totcogsusd.Sum().ToString("#,###");

        //    ViewBag.totalcogsidr = temptotcogsidr + ",00";
        //    ViewBag.totalcogsusd = temptotcogsusd + ",00";
        //    ViewBag.totpriceidr = temptotpriceidr + ",00";
        //    ViewBag.totpriceusd = temptotpriceusd + ",00";

        //    if (1 == 1)
        //    {
        //       return View(preprocGI);
        //    }
        //    else
        //    {
        //       return View(preprocGI);
        //    }
        //}

        [HttpPost]
        public async Task<List<VwGetNkpname>> ChangeListNKP(string vendor)
        {
            getUserNow();
            var data = await _context.VwGetNkpnames.Where(x => x.VendorId.Equals(vendor)).OrderBy(x => x.Nkpname).ToListAsync();
            
            ViewBag.itemdescs = data;
            return data;
        }

        [HttpPost]
        public async Task<List<VwGetVendor>> GetVendor()
        {
            getUserNow();
            var data = await _context.VwGetVendors.OrderBy(x=>x.VendorName).ToListAsync();

            return data;
        }

        [HttpPost]
        public async Task<List<MsCustomerVw>> GetListCustomer()
        {
            getUserNow();
            var data = await _context.MsCustomerVws.OrderBy(x => x.AccountName).ToListAsync();

            return data;
        }

        [HttpPost]
        public List<PreProcDetail> ReadExcel(IFormFile fileUpload, string preprocid)
        {
            getUserNow();
            var data = new List<PreProcDetail>();
            if (fileUpload != null)
            {
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/UploadExcel");

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                PreProcDetail tempDetail;

                string fileNameWithPath = Path.Combine(path, DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             +
                                                             fileUpload.FileName);

                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    fileUpload.CopyTo(stream);
                }

                var data2 = ImportExcel<PreProcFieldExcel>(fileNameWithPath, "Item Selected");

                foreach (var item in data2)
                {
                    tempDetail = new PreProcDetail();

                    tempDetail.Id = Int32.Parse(preprocid);
                    tempDetail.Principle = item.Principle;
                    tempDetail.Nkpcode = item.Nkpcode;
                    tempDetail.ItemDesc = item.ItemDesc;
                    tempDetail.Insurance = item.Insurance;
                    tempDetail.Warranty = item.Warranty;
                    tempDetail.Qty = item.Qty;
                    tempDetail.SurDukCheck = item.SurDukCheck;
                    tempDetail.Idnew = item.Id.ToString();
                    //_context.Add(tempDetail);
                    //await _context.SaveChangesAsync();
                    data.Add(tempDetail);
                }
            }
            return data;
        }
        [HttpPost]
        public async Task<TempPreProcDetail> SaveorUpdateDetail(string preprocid, int detailid, string no, double qty, string principal, string nkpcode, string itemdesc, bool insurance, bool surduk, int warranty, string presalesrev, string amrev, string risk, string commentam, string commentpresales)
        {
            var usernow = getUserNow();
            var data = new TempPreProcDetail();
            //if (fileUpload != null)
            //{
            //    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/UploadExcel");

            //    //create folder if not exist
            //    if (!Directory.Exists(path))
            //        Directory.CreateDirectory(path);

            //    TempPreProcDetail tempDetail;

            //    string fileNameWithPath = Path.Combine(path, DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
            //                                                 +
            //                                                 fileUpload.FileName);

            //    using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
            //    {
            //        fileUpload.CopyTo(stream);
            //    }

            //    var data2 = ImportExcel<PreProcFieldExcel>(fileNameWithPath, "Item Selected");

            //    foreach (var item in data2)
            //    {
            //        tempDetail = new TempPreProcDetail();

            //        tempDetail.Id = Int32.Parse(preprocid);
            //        tempDetail.Principle = item.Principle;
            //        tempDetail.Nkpcode = item.Nkpcode;
            //        tempDetail.ItemDesc = item.ItemDesc;
            //        tempDetail.Insurance = item.Insurance;
            //        tempDetail.Warranty = item.Warranty;
            //        tempDetail.Qty = item.Qty;
            //        tempDetail.SurDukCheck = item.SurDukCheck;
            //        tempDetail.Idnew = item.Id.ToString();
            //        _context.Add(tempDetail);
            //        await _context.SaveChangesAsync();
            //    }
            //    data = tempDetail;
            //}
            //else
            //{
            if (detailid == 0)
            {
                TempPreProcDetail tempDetail = new TempPreProcDetail();
                tempDetail.Id = Int32.Parse(preprocid);
                tempDetail.Principle = principal;
                tempDetail.Nkpcode = nkpcode;
                tempDetail.ItemDesc = itemdesc;
                tempDetail.Insurance = insurance;
                tempDetail.Warranty = warranty;
                tempDetail.Qty = qty;
                tempDetail.SurDukCheck = surduk;
                tempDetail.Idnew = no;
               

                _context.Add(tempDetail);
                await _context.SaveChangesAsync();
                data = tempDetail;
            }
            else
            {
                var tempDetail = _context.TempPreProcDetails.FirstOrDefault(s => s.DetailId.Equals(detailid));

                if (tempDetail != null)
                {
                    tempDetail.Id = Int32.Parse(preprocid);
                    tempDetail.Principle = principal;
                    tempDetail.Nkpcode = nkpcode;
                    tempDetail.ItemDesc = itemdesc;
                    tempDetail.Insurance = insurance;
                    tempDetail.Warranty = warranty;
                    tempDetail.Qty = qty;
                    tempDetail.SurDukCheck = surduk;
                    tempDetail.Idnew = no;
                    tempDetail.Risk = risk;
                    tempDetail.PresalesReview = presalesrev;
                    tempDetail.AmmanagerApproval = amrev;

                    PreProcApprovalHistory approvalhistory = new PreProcApprovalHistory();

                    approvalhistory.ApprovalDate = DateTime.Now;
                    approvalhistory.Remark = commentpresales;
                    approvalhistory.Status = "temp";
                    approvalhistory.DetailId = detailid;
                    approvalhistory.DomainName = User.Identity.Name;
                    approvalhistory.ApproverName = usernow;
                    approvalhistory.ApproverRole = "Presales";
                    approvalhistory.PreProcId = Int32.Parse(preprocid);

                    _context.Add(approvalhistory);
                }
                await _context.SaveChangesAsync();
                data = tempDetail;
            }
            //}

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> PostOverwriteOrAppendItem(string preprocid, double[] qty, string[] nkpcode, string[] itemdesc, bool[] sub, string[] headerno, string[] curr, double[] unitcogs, double[] unitrevenue , double[] totcogs, double[] totrevenue, string[] pid, string[] pidfull)
        {
            getUserNow();
            var data = new List<PreProcDetail>();
            PreProcDetail tempDetail;

            var count = 1;

            for (int i = 0; i < qty.Length; i++)
            {
                var principle = await _context.MsNkproducts.FirstOrDefaultAsync(x => x.NkpcodeLinkTitle == nkpcode[i]);

                tempDetail = new PreProcDetail();
                tempDetail.Id = Int32.Parse(preprocid);
                if (principle != null)
                {
                    tempDetail.Principle = principle.VendorId;
                    tempDetail.ItemDesc = principle.Nkpname;
                }
                tempDetail.Nkpcode = nkpcode[i];
                tempDetail.RemarkOnItem = itemdesc[i];
                tempDetail.Qty = qty[i];
                tempDetail.Idnew = count.ToString();
                tempDetail.SecondId = headerno[i] != null ? Int32.Parse(headerno[i]) : null;
                tempDetail.Currency = curr[i];
                tempDetail.UnitCogs = unitcogs[i];
                tempDetail.UnitPrice = unitrevenue[i];
                tempDetail.TotalCogs = totcogs[i].ToString();
                tempDetail.TotalPrice = totrevenue[i].ToString();
                tempDetail.Pidfull = pidfull[i];

                _context.Add(tempDetail);
                await _context.SaveChangesAsync();

                data.Add(tempDetail);
                count++;
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcFileAttachment>> GetAttachmentDetails (int procid, string name, int detailid)
        {
            getUserNow();
            var data = new List<PreProcFileAttachment>();

            if (name == "temp")
            {
                var specattach = await _context.PreProcFileAttachments.Where(x => x.PreProcId == procid && x.DocumentTypeId == 11 && x.Bastid == detailid).ToListAsync();
                var surdukattach = await _context.PreProcFileAttachments.Where(x => x.PreProcId == procid && x.DocumentTypeId == 12 && x.Bastid == detailid).ToListAsync();
                var vendorquoattach = await _context.PreProcFileAttachments.Where(x => x.PreProcId == procid && x.DocumentTypeId == 13 && x.Bastid == detailid).ToListAsync();

                if (specattach != null)
                {
                    foreach (var item in specattach)
                    {
                       data.Add(item);
                    }
                }
                if (surdukattach != null)
                {
                    foreach (var item in surdukattach)
                    {
                        data.Add(item);
                    }
                }
                if (vendorquoattach != null)
                {
                    foreach (var item in vendorquoattach)
                    {
                        data.Add(item);
                    }
                }
            }
            else
            {
                var detail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid);
                if (detail != null)
                {
                    //    var specattach = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == detail.SpecAttachId);
                    //    var surdukattach = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == detail.SurdukAttachId);

                    //    if (specattach != null && surdukattach != null)
                    //    {
                    //        data.Add(specattach);
                    //        data.Add(surdukattach);
                    //    }
                    //    else if (specattach != null)
                    //    {
                    //        data.Add(specattach);
                    //    }
                    //    else if (surdukattach != null)
                    //    {
                    //        data.Add(surdukattach);
                    //    }

                    var specattach = await _context.PreProcFileAttachments.Where(x => x.PreProcId == procid && x.DocumentTypeId == 11 && x.Bastid == detail.DetailId).ToListAsync();
                    var surdukattach = await _context.PreProcFileAttachments.Where(x => x.PreProcId == procid && x.DocumentTypeId == 12 && x.Bastid == detail.DetailId).ToListAsync();
                    var vendorquoattach = await _context.PreProcFileAttachments.Where(x => x.PreProcId == procid && x.DocumentTypeId == 13 && x.Bastid == detail.DetailId).ToListAsync();

                    if (specattach != null)
                    {
                        foreach (var item in specattach)
                        {
                            PreProcFileAttachment tempfile = new PreProcFileAttachment();

                            tempfile.Id = item.Id;
                            tempfile.StrictedFileName = item.StrictedFileName;
                            tempfile.FileName = item.FileName;
                            tempfile.DocumentTypeId = item.DocumentTypeId;

                            data.Add(tempfile);
                        }
                    }

                    if (vendorquoattach != null)
                    {
                        foreach (var item in vendorquoattach)
                        {
                            PreProcFileAttachment tempfile = new PreProcFileAttachment();

                            tempfile.Id = item.Id;
                            tempfile.StrictedFileName = item.StrictedFileName;
                            tempfile.FileName = item.FileName;
                            tempfile.DocumentTypeId = item.DocumentTypeId;

                            data.Add(tempfile);
                        }
                    }
                    if (surdukattach != null)
                    {
                        foreach (var item in surdukattach)
                        {
                            PreProcFileAttachment tempfile = new PreProcFileAttachment();

                            tempfile.Id = item.Id;
                            tempfile.StrictedFileName = item.StrictedFileName;
                            tempfile.FileName = item.FileName;
                            tempfile.DocumentTypeId = item.DocumentTypeId;

                            data.Add(tempfile);
                        }
                    }
                }
            }
            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> GetItemCheckedDetails(int procid, int[] detailid)
        {
            getUserNow();
            var data = new List<PreProcDetail>();

            if (detailid != null)
            {
                foreach (var item in detailid)
                {
                    var tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == item && x.Id == procid);
                    data.Add(tempData);
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<PreProcPmochecklist> PostSubmitPMOChecklist(PreProcGeneralInfo2 model, int procid, string custpic, string highlvlsow, string projectsite, string eststartdate, string enddate, string estkickoff, string kickoffdate, string prodteampic, IFormFile filespk, IFormFile filerfp, IFormFile filequot, IFormFile fileproposal, IFormFile fileboq, IFormFile filenego, IFormFile fileother, string specterm, double projectval,bool checkCustPIC,bool checkHighLvlSOW,bool checkProjectSite,bool checkProjectValue,bool checkEstStartDate,bool checkEstKickOff,bool checkProductTeamPIC,bool checkAttachSPK,bool checkAttachRFP,bool checkAttachQuot,bool checkAttachProposal,bool checkAttachBOQ,bool checkAttachNego,bool checkAttachOther, int? pmmandays, int? pcmandays, decimal? pmvalue, decimal? pcvalue)
        {
            getUserNow();
            var data = new PreProcPmochecklist();

            var tempData = await _context.PreProcPmochecklists.FirstOrDefaultAsync(x => x.PreProcId == procid);

            var sendemail = false;

            var preproc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);

            PreProcHeaderAttachment preProcAttach;

            if (tempData == null)
            {
                PreProcPmochecklist pmo = new PreProcPmochecklist();
                pmo.PreProcId = procid;
                pmo.CustomerPic = custpic;
                pmo.HighLevelSow = highlvlsow;
                pmo.ProjectSite = projectsite;
                pmo.ProjectValue = projectval.ToString();
                pmo.EstimatedStartDate = eststartdate != null ? DateTime.Parse(eststartdate) : null;
                pmo.EndDate = eststartdate != null ? DateTime.Parse(enddate) : null;
                pmo.EstimatedKickOffNote = estkickoff;
                pmo.EstimatedKickOffDate = kickoffdate == null ? null : DateTime.Parse(kickoffdate);
                pmo.ProductTeamPic = prodteampic;
                pmo.SpecificTerms = specterm;
                pmo.CheckedCustomerPic = checkCustPIC;
                pmo.CheckedHighLevelSow = checkHighLvlSOW;
                pmo.CheckedProjectSite = checkProjectSite;
                pmo.CheckedProjectValue = checkProjectValue;
                pmo.CheckedEstimatedStartEndDate = checkEstStartDate;
                pmo.CheckedEstimatedKickOff = checkEstKickOff;
                pmo.CheckedProductTeamPic = checkProductTeamPIC;
                pmo.CheckedSpkorPoorContract = checkAttachSPK;
                pmo.CheckedRfporRksorTor = checkAttachRFP;
                pmo.CheckedQuotation = checkAttachQuot;
                pmo.CheckedProposal = checkAttachProposal;
                pmo.CheckedBoQ = checkAttachBOQ;
                pmo.CheckedNegotiationMoM = checkAttachNego;
                pmo.CheckedOtherAttachment = checkAttachOther;
                pmo.Pmmandays = pmmandays;
                pmo.Pcmandays = pcmandays;
                pmo.Pmvalue = pmvalue;
                pmo.Pcvalue = pcvalue;

                _context.Add(pmo);
                await _context.SaveChangesAsync();

                preproc.Cvpc = model.Cvpc;
                preproc.Cvpm = model.Cvpm;
                await _context.SaveChangesAsync();

                //data = pmo;
            }
            else
            {
                tempData.PreProcId = procid;
                tempData.CustomerPic = custpic;
                tempData.HighLevelSow = highlvlsow;
                tempData.ProjectSite = projectsite;
                tempData.ProjectValue = projectval.ToString();
                tempData.EstimatedStartDate = DateTime.Parse(eststartdate);
                tempData.EndDate = DateTime.Parse(enddate);
                tempData.EstimatedKickOffNote = estkickoff;
                tempData.EstimatedKickOffDate = kickoffdate == null ? null : DateTime.Parse(kickoffdate);
                tempData.ProductTeamPic = prodteampic;
                tempData.SpecificTerms = specterm;
                tempData.CheckedCustomerPic = checkCustPIC;
                tempData.CheckedHighLevelSow = checkHighLvlSOW;
                tempData.CheckedProjectSite = checkProjectSite;
                tempData.CheckedProjectValue = checkProjectValue;
                tempData.CheckedEstimatedStartEndDate = checkEstStartDate;
                tempData.CheckedEstimatedKickOff = checkEstKickOff;
                tempData.CheckedProductTeamPic = checkProductTeamPIC;
                tempData.CheckedSpkorPoorContract = checkAttachSPK;
                tempData.CheckedRfporRksorTor = checkAttachRFP;
                tempData.CheckedQuotation = checkAttachQuot;
                tempData.CheckedProposal = checkAttachProposal;
                tempData.CheckedBoQ = checkAttachBOQ;
                tempData.CheckedNegotiationMoM = checkAttachNego;
                tempData.CheckedOtherAttachment = checkAttachOther;

                await _context.SaveChangesAsync();

                //data = tempData;
                sendemail = true;
            }

            if (filespk != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                filespk.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("SPK")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    filespk.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = filespk.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (filerfp != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                filerfp.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("RFP")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    filerfp.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = filerfp.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (fileboq != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                fileboq.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("BOQ")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    fileboq.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = fileboq.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (filequot != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                filequot.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("QUO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    filequot.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = filequot.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (filenego != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                filenego.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("NMM")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    filenego.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = filenego.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (fileproposal != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                fileproposal.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("PRO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    fileproposal.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = fileproposal.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (fileother != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                fileother.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("OTH")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    fileother.CopyTo(stream);
                }
                preProcAttach.Folder = preproc.PresalesId;
                preProcAttach.FileName = fileother.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preproc.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }

            if (sendemail)
            {
                await SendNotifToPMO(procid);
            }

            return data;
        }

        [HttpPost]
        public async Task<PreProcRiskAssesmentDatum> PostRiskFactorAssess(int procid, string selectOverallRisk, string txtRiskDesc, string selectSourceBudget, string selectCompetition, string selectContingency, string txtContingencyCost, string txtCompetitor, string txtEstProjectStart, string txtEstProjectEnd, string selectProcType, string selectProjectSource, string txtRevMLPT, string txtTotSubCon, string txtProcCondition, string txtPredefinedCondition, string resultSelectedCommitteeDomain, string[] selectSubconName, string[] txtSubconName, string[] txtPercent, string[] txtHighLvlScope, string[] txtPreload, string[] checkFactors)
        {
            var curruser = getUserNow();
            var data = new PreProcRiskAssesmentDatum();
            var riskassess = await _context.PreProcRiskAssesmentData.FirstOrDefaultAsync(x => x.PreProcId == procid);
            var preproc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);
            var temp = new PreProcRiskAssesmentDatum();
            string overallrisk = "";
            var listriskdata = new List<PreProcRiskDatum>();
            var listprincipalprojrev = new List<PreProcPrincipalProjectRevenue>();

            if (preproc != null)
            {
                overallrisk = selectOverallRisk != preproc.OverallRisk ? selectOverallRisk : "";
                preproc.OverallRisk = selectOverallRisk;
                await _context.SaveChangesAsync();
            }

            if (riskassess == null)
            {
                PreProcRiskAssesmentDatum rad = new PreProcRiskAssesmentDatum();
                rad.PreProcId = procid;
                rad.Description = txtRiskDesc;
                rad.SourceBudget = selectSourceBudget;
                rad.EstProjectStart = txtEstProjectStart != null ? DateTime.Parse(txtEstProjectStart) : null;
                rad.EstProjectCompletion = txtEstProjectStart != null ? DateTime.Parse(txtEstProjectEnd) : null;
                rad.ProcurementType = selectProcType;
                rad.Competition = selectCompetition;
                rad.ProjectSource = selectProjectSource;
                rad.Competitor = txtCompetitor;
                rad.ContigencyCost = txtContingencyCost != null && selectContingency == "True" ? decimal.Parse(txtContingencyCost) : null;
                rad.PredefinedConditionReqByPrincipal = txtPredefinedCondition;
                rad.ProjectCondition = txtProcCondition;
                if (resultSelectedCommitteeDomain != null)
                {
                    rad.EmailAssignTo = resultSelectedCommitteeDomain;
                }
                else
                {
                    var manageram = await _context.VwEmployeeBases.FirstOrDefaultAsync(x => x.DomainName == User.Identity.Name);
                    if (manageram != null)
                    {
                       rad.EmailAssignTo = manageram.Manager1Email + ";";
                    }
                }

                _context.Add(rad);
                await _context.SaveChangesAsync();

                PreProcPrincipalProjectRevenue newprojectrev;

                if (txtRevMLPT != null || txtRevMLPT != "")
                {
                    newprojectrev = new PreProcPrincipalProjectRevenue();

                    newprojectrev.PreProcId = procid;
                    newprojectrev.PrincipalName = "MLPT";
                    newprojectrev.ProjectRevenue = txtRevMLPT != null ? decimal.Parse(txtRevMLPT) : null;

                    _context.Add(newprojectrev);
                    await _context.SaveChangesAsync();
                }

                for (int i = 0; i < txtPercent.Length; i++)
                {
                    newprojectrev = new PreProcPrincipalProjectRevenue();

                    newprojectrev.PreProcId = procid;
                    newprojectrev.PrincipalName = selectSubconName[i] == "Others" ? txtSubconName[i] : selectSubconName[i];
                    newprojectrev.ProjectRevenue = decimal.Parse(txtPercent[i]);
                    newprojectrev.PrincipalScope = txtHighLvlScope[i];
                    newprojectrev.Preload = Int32.Parse(txtPreload[i]);

                    _context.Add(newprojectrev);
                    await _context.SaveChangesAsync();
                }
                PreProcRiskDatum newriskdata;

                for (int i = 0; i < checkFactors.Length; i++)
                {
                    var tempsplit = checkFactors[i].Split(";");
                    var risk = await _context.PreProcRisks.FirstOrDefaultAsync(x => x.RiskName.ToLower().Replace(" ", "").Equals(tempsplit[0].Replace(" ", "")));

                    if (tempsplit[1] == "true")
                    {
                        newriskdata = new PreProcRiskDatum();

                        newriskdata.PreProcId = procid;
                        newriskdata.RiskOrder = risk.RiskOrderNo.Value;
                        newriskdata.CheckedBy = "stepanus.triatmaja@multipolar.com";
                        newriskdata.RiskType = risk.RiskType;
                        newriskdata.RiskId = risk.RiskId;

                        _context.Add(newriskdata);
                        await _context.SaveChangesAsync();
                    }
                    
                }
                data = rad;

                await SendNotifByType("", procid, "NewItem", new List<PreProcPodatum>(), new List<PreProcDetail>(), data, "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());

            }
            else
            {
                temp.Description = riskassess.Description;
                temp.SourceBudget = riskassess.SourceBudget;
                temp.EstProjectStart = riskassess.EstProjectStart;
                temp.EstProjectCompletion = riskassess.EstProjectCompletion;
                temp.ProcurementType = riskassess.ProcurementType;
                temp.Competition = riskassess.Competition;
                temp.ProjectSource = riskassess.ProjectSource;
                temp.Competitor = riskassess.Competitor;
                temp.ContigencyCost = riskassess.ContigencyCost;
                temp.PredefinedConditionReqByPrincipal = riskassess.PredefinedConditionReqByPrincipal;
                temp.ProjectCondition = riskassess.ProjectCondition;
                temp.EmailAssignTo = riskassess.EmailAssignTo;

                riskassess.PreProcId = procid;
                riskassess.Description = txtRiskDesc;
                riskassess.SourceBudget = selectSourceBudget;
                riskassess.EstProjectStart = txtEstProjectStart != null ? DateTime.Parse(txtEstProjectStart) : null;
                riskassess.EstProjectCompletion = txtEstProjectStart != null ? DateTime.Parse(txtEstProjectEnd) : null;
                riskassess.ProcurementType = selectProcType;
                riskassess.Competition = selectCompetition;
                riskassess.ProjectSource = selectProjectSource;
                riskassess.Competitor = txtCompetitor;
                riskassess.ContigencyCost = txtContingencyCost != null && selectContingency == "True" ? decimal.Parse(txtContingencyCost) : null;
                riskassess.PredefinedConditionReqByPrincipal = txtPredefinedCondition;
                riskassess.ProjectCondition = txtProcCondition;
                riskassess.EmailAssignTo = resultSelectedCommitteeDomain;

                await _context.SaveChangesAsync();

                PreProcPrincipalProjectRevenue newprojectrev;

                if (txtRevMLPT != null || txtRevMLPT != "")
                {
                    var projectrev = await _context.PreProcPrincipalProjectRevenues.FirstOrDefaultAsync(x => x.PreProcId == procid && x.PrincipalName.Equals("MLPT"));

                    if (projectrev != null)
                    {
                        var tempprojrev = projectrev.ProjectRevenue;

                        projectrev.ProjectRevenue = txtRevMLPT != null ? decimal.Parse(txtRevMLPT) : null;

                        await _context.SaveChangesAsync();

                        
                        if (txtRevMLPT != null)
                        {
                            var tempprincipalprojrev = new PreProcPrincipalProjectRevenue();
                            tempprojrev = tempprojrev != decimal.Parse(txtRevMLPT) ? decimal.Parse(txtRevMLPT) : null;
                            tempprincipalprojrev.ProjectRevenue = tempprojrev;
                            tempprincipalprojrev.PrincipalName = "MLPT";

                            listprincipalprojrev.Add(tempprincipalprojrev);
                        }
                    }
                    else
                    {
                        newprojectrev = new PreProcPrincipalProjectRevenue();

                        newprojectrev.PreProcId = procid;
                        newprojectrev.PrincipalName = "MLPT";
                        newprojectrev.ProjectRevenue = txtRevMLPT != null ? decimal.Parse(txtRevMLPT) : null;

                        _context.Add(newprojectrev);
                        await _context.SaveChangesAsync();

                        listprincipalprojrev.Add(newprojectrev);
                    }
                }

                var projectrevsub = await _context.PreProcPrincipalProjectRevenues.Where(x => x.PreProcId == procid && !x.PrincipalName.Equals("MLPT")).ToListAsync();
                var countprojrevsub = projectrevsub.Count;
                _context.RemoveRange(projectrevsub);
                await _context.SaveChangesAsync();

                var tempprincipalproj = new List<PreProcPrincipalProjectRevenue>();

                for (int i = 0; i < txtPercent.Length; i++)
                {
                    newprojectrev = new PreProcPrincipalProjectRevenue();

                    newprojectrev.PreProcId = procid;
                    newprojectrev.PrincipalName = selectSubconName[i] == "Others" ? txtSubconName[i] : selectSubconName[i];
                    newprojectrev.ProjectRevenue = txtPercent[i] != null ? decimal.Parse(txtPercent[i]) : null;
                    newprojectrev.PrincipalScope = txtHighLvlScope[i];
                    newprojectrev.Preload = txtPreload[i] != null ? Int32.Parse(txtPreload[i]) : null;

                    _context.Add(newprojectrev);
                    await _context.SaveChangesAsync();

                    tempprincipalproj.Add(newprojectrev);
                }
                
                if (tempprincipalproj.Count != countprojrevsub)
                {
                    listprincipalprojrev.AddRange(tempprincipalproj);
                }
                else if (listprincipalprojrev.Count > 0 && tempprincipalproj.Count > 0)
                {
                    listprincipalprojrev.AddRange(tempprincipalproj);
                }

                PreProcRiskDatum newriskdata;

                for (int i = 0; i < checkFactors.Length; i++)
                {
                    var tempsplit = checkFactors[i].Split(";");
                    var risk = await _context.PreProcRisks.FirstOrDefaultAsync(x => x.RiskName.ToLower().Replace(" ", "").Equals(tempsplit[0].Replace(" ", "")));
                    var riskdata = await _context.PreProcRiskData.FirstOrDefaultAsync(x => x.PreProcId == procid && x.RiskId == risk.RiskId && !x.CheckedBy.Contains("UncheckedBy"));

                    if (tempsplit[1] == "true")
                    {
                        if (riskdata == null)
                        {
                            newriskdata = new PreProcRiskDatum();

                            newriskdata.PreProcId = procid;
                            newriskdata.RiskOrder = risk.RiskOrderNo.Value;
                            newriskdata.CheckedBy = User.Identity.Name;
                            newriskdata.RiskType = risk.RiskType;
                            newriskdata.RiskId = risk.RiskId;

                            _context.Add(newriskdata);
                            await _context.SaveChangesAsync();

                            listriskdata.Add(newriskdata);
                        }
                    }
                    else
                    {
                        if (riskdata != null)
                        {
                            riskdata.CheckedBy = "UncheckedBy;" + User.Identity.Name;

                            await _context.SaveChangesAsync();
                        }
                    }
                }
                data = riskassess;

                var temprisk = await _context.PreProcRiskAssesmentData.FirstOrDefaultAsync(x => x.PreProcId == procid);

                temprisk.Description = riskassess.Description != temp.Description ? riskassess.Description : null;
                temprisk.SourceBudget = riskassess.SourceBudget != temp.SourceBudget ? riskassess.SourceBudget : null;
                temprisk.EstProjectStart = riskassess.EstProjectStart != temp.EstProjectStart ? riskassess.EstProjectStart : null;
                temprisk.EstProjectCompletion = riskassess.EstProjectCompletion != temp.EstProjectCompletion ? riskassess.EstProjectCompletion : null;
                temprisk.ProcurementType = riskassess.ProcurementType != temp.ProcurementType ? riskassess.ProcurementType : null;
                temprisk.Competition = riskassess.Competition != temp.Competition ? riskassess.Competition : null;
                temprisk.ProjectSource = riskassess.ProjectSource != temp.ProjectSource ? riskassess.ProjectSource : null;
                temprisk.Competitor = riskassess.Competitor != temp.Competitor ? riskassess.Competitor : null;
                temprisk.ContigencyCost = riskassess.ContigencyCost != temp.ContigencyCost ? riskassess.ContigencyCost : null;
                temprisk.PredefinedConditionReqByPrincipal = riskassess.PredefinedConditionReqByPrincipal != temp.PredefinedConditionReqByPrincipal ? riskassess.PredefinedConditionReqByPrincipal : null;
                temprisk.ProjectCondition = riskassess.ProjectCondition != temp.ProjectCondition ? riskassess.ProjectCondition : null;
                temprisk.EmailAssignTo = riskassess.EmailAssignTo != temp.EmailAssignTo ? riskassess.EmailAssignTo : null;

                await SendNotifByType("", procid, "UpdatedRiskAssess", new List<PreProcPodatum>(), new List<PreProcDetail>(), temprisk, overallrisk, listriskdata, listprincipalprojrev);
            }


            return data;
        }

        [HttpPost]
        public async Task<PreProcRiskAssesmentDatum> PostReviewRiskFactor(int procid, string reviewrisk)
        {
            var usernow = getUserNow();
            var data = new PreProcRiskAssesmentDatum();
            PreProcRiskAssesmentDatum riskAssesmentDatum;

            var riskassess = await _context.PreProcRiskAssesmentData.FirstOrDefaultAsync(x => x.PreProcId == procid);

            if (riskassess != null)
            {
                var preproc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);

                if (preproc != null)
                {
                    preproc.InitialRiskReview = true;
                    await _context.SaveChangesAsync();
                }

                riskassess.ReviewDate = DateTime.Now;
                await _context.SaveChangesAsync();

                var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Risk");

                PreProcApprovalHistory approvalhistory = new PreProcApprovalHistory();

                approvalhistory.Remark = reviewrisk;
                approvalhistory.Status = "Reviewed";
                approvalhistory.DomainName = User.Identity.Name;
                approvalhistory.ApproverName = usernow;
                approvalhistory.ApproverRole = "Committee";
                approvalhistory.ApprovalDate = DateTime.Now;
                approvalhistory.Type = commenttype.Id;
                approvalhistory.PreProcId = procid;

                _context.Add(approvalhistory);
                await _context.SaveChangesAsync();
            }

            return data;
        }

        [HttpPost]
        public async Task<PreProcApprovalHistory> PostApprovalRiskAssessment(int procid, string commentapproval, string status)
        {
            var usernow = getUserNow();
            var data = new PreProcApprovalHistory();
            PreProcApprovalHistory preProcApproval;

            var preproc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);
            var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Risk");
            var rand = new Random(); 
            var uid = rand.Next(1, 5);

            if (status == "Approved")
            {
                PreProcApprovalHistory approvalhistory = new PreProcApprovalHistory();

                approvalhistory.Remark = commentapproval;
                approvalhistory.Status = status;
                approvalhistory.DomainName = User.Identity.Name;
                approvalhistory.ApproverName = usernow;
                approvalhistory.ApproverRole = "BOD";
                approvalhistory.ApprovalDate = DateTime.Now;
                approvalhistory.Type = commenttype.Id;
                approvalhistory.PreProcId = procid;

                _context.Add(approvalhistory);
                await _context.SaveChangesAsync();

            }
            else
            {
                PreProcApprovalHistory approvalhistory = new PreProcApprovalHistory();

                approvalhistory.Remark = commentapproval;
                approvalhistory.Status = status;
                approvalhistory.DomainName = User.Identity.Name;
                approvalhistory.ApproverName = usernow;
                approvalhistory.ApproverRole = "BOD";
                approvalhistory.ApprovalDate = DateTime.Now;
                approvalhistory.Type = commenttype.Id;
                approvalhistory.PreProcId = procid;

                _context.Add(approvalhistory);
                await _context.SaveChangesAsync();
            }

            if (preproc != null)
            {
                var countapproved = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == procid && x.Type == commenttype.Id && x.Status.Equals("Approved")).ToListAsync();
                var countrejected = await _context.PreProcApprovalHistories.Where(x => x.PreProcId == procid && x.Type == commenttype.Id && x.Status.Equals("Rejected")).ToListAsync();

                if (countapproved.Count == 3 && preproc.Stage == "RiskAssessment")
                {
                    preproc.Type = "Risk Approved";
                    preproc.Stage = "Presales";
                    await _context.SaveChangesAsync();
                }
                else if (countrejected.Count == 3 && preproc.Stage == "RiskAssessment")
                {
                    preproc.Type = "Risk Rejected";
                    preproc.Stage = "Presales";
                    await _context.SaveChangesAsync();
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> PostUpdateRevenue(int[] procid, int[] detailid, string[] currency, double[] unitcogs, double[] unitrevenue)
        {
            getUserNow();
            var data = new List<PreProcDetail>();

            if (detailid != null)
            {
                var tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[0] && x.Id == procid[0]);
                for (var i = 0; i < detailid.Length; i++)
                {
                    var temp = new PreProcDetail();
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[i] && x.Id == procid[0]);

                    temp.Currency = currency[i];
                    temp.UnitCogs = unitcogs[i];
                    temp.UnitPrice = unitrevenue[i];
                    temp.DetailId = detailid[i];
                    temp.TotalCogs = (unitcogs[i] * tempData.Qty).Value.ToString();
                    temp.TotalPrice = (unitrevenue[i] * tempData.Qty).Value.ToString();

                    data.Add(temp);
                }

                foreach (var item in data)
                {
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == item.DetailId && x.Id == procid[0]);

                    if (tempData != null)
                    {
                        tempData.Currency = item.Currency;
                        tempData.UnitCogs = item.UnitCogs;
                        tempData.TotalCogs = (item.UnitCogs * tempData.Qty).Value.ToString();
                        tempData.UnitPrice = item.UnitPrice;
                        tempData.TotalPrice = (item.UnitPrice * tempData.Qty).Value.ToString();

                        await _context.SaveChangesAsync();
                        //data.Add(tempData);
                    }
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> PostUpdateCOGS(int[] procid, int[] detailid, string[] currency, double[] unitcogs)
        {
            getUserNow();
            var data = new List<PreProcDetail>();

            if (detailid != null)
            {
                var tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[0] && x.Id == procid[0]);
                for (var i = 0; i < detailid.Length; i++)
                {
                    var temp = new PreProcDetail();
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[i] && x.Id == procid[0]);

                    temp.Currency = currency[i];
                    temp.UnitCogs = unitcogs[i];
                    temp.DetailId = detailid[i];
                    temp.TotalCogs = (unitcogs[i] * tempData.Qty).Value.ToString();

                    data.Add(temp);
                }

                foreach (var item in data)
                {
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == item.DetailId && x.Id == procid[0]);

                    if (tempData != null)
                    {
                        tempData.Currency = item.Currency;
                        tempData.UnitCogs = item.UnitCogs;
                        tempData.TotalCogs = (item.UnitCogs * tempData.Qty).Value.ToString();

                        await _context.SaveChangesAsync();
                        //data.Add(tempData);
                    }
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> PostUpdateRisk(int[] procid, int[] detailid, string[] risk)
        {
            getUserNow();
            var data = new List<PreProcDetail>();

            if (detailid != null)
            {
                var tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[0] && x.Id == procid[0]);
                for (var i = 0; i < detailid.Length; i++)
                {
                    var temp = new PreProcDetail();
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[i] && x.Id == procid[0]);

                    temp.Risk = risk[i];
                    temp.DetailId = detailid[i];

                    data.Add(temp);
                }

                foreach (var item in data)
                {
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == item.DetailId && x.Id == procid[0]);

                    if (tempData != null)
                    {
                        tempData.Risk = item.Risk;

                        await _context.SaveChangesAsync();
                        //data.Add(tempData);
                    }
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcApprovalHistory>> PostUpdateReview(int[] procid, int[] detailid, string[] status, string[] comment)
        {
            var usernow = getUserNow();
            var curremail = getUserNowEmail();

            var paramemail = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Current Login Email").FirstOrDefaultAsync();

            if (paramemail.AlphaNumericValue != null && paramemail.AlphaNumericValue != "")
            {
                curremail = paramemail.AlphaNumericValue;
            }

            //var presales = await _context.VwGetPresalesMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail == curremail);
            var presales = await _context.VwGetPossibleProductManagers.FirstOrDefaultAsync(x => x.EmployeeEmail == curremail);
            var am = await _context.VwGetAmmembers.FirstOrDefaultAsync(x => x.EmployeeEmail == curremail && x.JobTitleLevel.Contains("Head"));
            var role = "";

            if (presales != null)
            {
                role = "Presales";
            }
            else if (am != null)
            {
                role = "AM Manager";
            }

            var data = new List<PreProcApprovalHistory>();
            PreProcApprovalHistory approvalhistory;
            if (detailid != null)
            {
                var tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[0] && x.Id == procid[0]);
                for (var i = 0; i < detailid.Length; i++)
                {
                    tempData = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[i] && x.Id == procid[0]);

                    if (tempData != null)
                    {
                        if (role == "Presales")
                        {
                            tempData.PresalesReview = status[i];

                            await _context.SaveChangesAsync();
                        }
                        else if (role == "AM Manager")
                        {
                            tempData.AmmanagerApproval = status[i];

                            await _context.SaveChangesAsync();
                        }

                        
                        //data.Add(tempData);
                    }
                    var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Detail");

                    var lineitem = await _context.PreProcDetails.Where(x => x.Id == procid[0]).ToListAsync();
                    var line = 0;
                    for (var j = 0; j < lineitem.Count; j++)
                    {
                        if (lineitem[j].DetailId == detailid[i])
                        {
                            line = j+1;
                            break;
                        }
                    }

                    approvalhistory = new PreProcApprovalHistory();

                    approvalhistory.ApprovalDate = DateTime.Now;
                    approvalhistory.Remark = comment[i];
                    approvalhistory.Status = status[i] + ": Line " + line;
                    approvalhistory.DetailId = detailid[i];
                    approvalhistory.DomainName = User.Identity.Name;
                    approvalhistory.ApproverName = usernow;
                    approvalhistory.ApproverRole = role;
                    approvalhistory.Type = commenttype.Id;
                    approvalhistory.PreProcId = Int32.Parse(procid[0].ToString());
                    
                    _context.Add(approvalhistory);
                    await _context.SaveChangesAsync();

                    var tempapproval = new PreProcApprovalHistory();

                    tempapproval.ApprovalDate = DateTime.Now;
                    tempapproval.Remark = comment[i];
                    tempapproval.Status = status[i];
                    tempapproval.DetailId = detailid[i];
                    tempapproval.DomainName = User.Identity.Name;
                    tempapproval.ApproverName = usernow;
                    tempapproval.ApproverRole = role;
                    tempapproval.PreProcId = Int32.Parse(procid[0].ToString());

                    data.Add(tempapproval);
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcPodatum>> PostAddPOData(int[] procid, string[] noitemcs, string[] pidnum)
        {
            getUserNow();
            var data = new List<PreProcPodatum>();
            PreProcPodatum temp;

            for (var i = 0; i < procid.Length; i++)
            {
                temp = new PreProcPodatum();
                var tempData = await _context.VwGetMplPoAlls.FirstOrDefaultAsync(x => x.Pidnum == pidnum[i]);
                var tempdetail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.Id == procid[i] && x.Idnew == noitemcs[i]);

                temp.PreProcId = procid[i];
                //if (tempdetail.DetailIdold != null)
                //{
                //    temp.DetailId = tempdetail.DetailIdold.Value;
                //}
                temp.DetailId = tempdetail.DetailId;
                temp.Poid = Int32.Parse(tempData.NoPo);
                temp.Podate = tempData.TanggalPo;
                temp.PolineId = pidnum[i];
                temp.Poqty = Int32.Parse(tempData.Quantity);
                temp.Idnew = noitemcs[i];
                temp.VendorName = tempData.NamaVendor;
                temp.Buyer = tempData.Buyer;
                temp.Term = tempData.Term;
                temp.ItemDescription = tempData.Description;
                temp.ItemNo = tempData.ItemNo;
                temp.Curr = tempData.CurrencyCode;
                temp.UnitPrice = tempData.UnitPrice == null || tempData.UnitPrice == "" ? null : decimal.Parse(tempData.UnitPrice);
                temp.Uom = tempData.Uom;
                temp.VendorCatalog = tempData.VendorCatalog;

                _context.Add(temp);
                await _context.SaveChangesAsync();
                
                data.Add(temp);

            }

            //foreach (var item in data)
            //{
            //    tempData = await _context.TempPreProcDetails.FirstOrDefaultAsync(x => x.DetailId == item.DetailId && x.Id == procid[0]);

            //    if (tempData != null)
            //    {
            //        tempData.ProductManager = item.ProductManager;

            //        await _context.SaveChangesAsync();
            //        //data.Add(tempData);
            //    }
            //}

            await SendNotifByType("", data[0].PreProcId, "UpdatedPOFromProc", data, new List<PreProcDetail>(), new PreProcRiskAssesmentDatum(), "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcPodatum>> PostUpdatePOData(int procid, string[] idnew, string[] poline, string inputtype, string txtExpPODateSingle, string txtETDSingle, string txtPOQtySingle, string txtETASingle, string[] txtExpPODateMultiple, string[] txtETDMultiple, string[] txtPOQtyMultiple, string[] txtETAMultiple)
        {
            getUserNow();
            var data = new List<PreProcPodatum>();
            var passdata = new List<PreProcPodatum>();
            //PreProcPodatum temp;
            var temp = new List<PreProcPodatum>();
            
            if (inputtype == "Single")
            {
                for (var i = 0; i < idnew.Length; i++)
                {
                    var tempdetail = await _context.PreProcPodata.FirstOrDefaultAsync(x => x.PreProcId == procid && x.PolineId == poline[i] && x.Idnew == idnew[i]);
                    var temppassdata = new PreProcPodatum();
                    var qty = tempdetail.Poqty;
                    var date = tempdetail.Eta;
                    var etd = tempdetail.Etd;
                    var podate = tempdetail.ExpectedPoissuedDate;

                    if (tempdetail != null)
                    {
                        tempdetail.ExpectedPoissuedDate = txtExpPODateSingle != null && txtExpPODateSingle != "undefined" ? DateTime.Parse(txtExpPODateSingle) : tempdetail.ExpectedPoissuedDate;
                        tempdetail.Poqty = txtPOQtySingle != null && txtPOQtySingle != "0" && txtPOQtySingle != "undefined" ? Int32.Parse(txtPOQtySingle) : tempdetail.Poqty;
                        tempdetail.Eta = txtETASingle != null && txtETASingle != "undefined" ? DateTime.Parse(txtETASingle) : tempdetail.Eta;
                        tempdetail.Etd = txtETDSingle != null && txtETDSingle != "undefined" ? DateTime.Parse(txtETDSingle) : tempdetail.Etd;
                        
                        await _context.SaveChangesAsync();

                        temppassdata.Eta = date;
                        temppassdata.Poqty = qty;
                        temppassdata.Etd = etd;
                        temppassdata.ExpectedPoissuedDate = podate;
                    }
                    temp.Add(temppassdata);
                    data.Add(tempdetail);
                }
                for (var i = 0; i < idnew.Length; i++)
                {
                    temp[i].Eta = temp[i].Eta != data[i].Eta ? data[i].Eta : null;
                    temp[i].Poqty = temp[i].Poqty != data[i].Poqty ? data[i].Poqty : null;
                    temp[i].ExpectedPoissuedDate = temp[i].ExpectedPoissuedDate != data[i].ExpectedPoissuedDate ? data[i].ExpectedPoissuedDate : null;
                    temp[i].Etd = temp[i].Etd != data[i].Etd ? data[i].Etd : null;

                    passdata.Add(data[i]);
                    passdata[i].Eta = temp[i].Eta;
                    passdata[i].Poqty = temp[i].Poqty;
                    passdata[i].ExpectedPoissuedDate = temp[i].ExpectedPoissuedDate;
                    passdata[i].Etd = temp[i].Etd;
                }
            }
            else
            {
                for (var i = 0; i < idnew.Length; i++)
                {
                    var tempdetail = await _context.PreProcPodata.FirstOrDefaultAsync(x => x.PreProcId == procid && x.PolineId == poline[i] && x.Idnew == idnew[i]);
                    var temppassdata = new PreProcPodatum();
                    var qty = tempdetail.Poqty;
                    var date = tempdetail.Eta;
                    var etd = tempdetail.Etd;
                    var podate = tempdetail.ExpectedPoissuedDate;

                    if (tempdetail != null)
                    {
                        tempdetail.ExpectedPoissuedDate = txtExpPODateMultiple[i] != null && txtExpPODateMultiple[i] != "undefined" ? DateTime.Parse(txtExpPODateMultiple[i]) : tempdetail.ExpectedPoissuedDate;
                        tempdetail.Poqty = txtPOQtyMultiple[i]!= null && txtPOQtyMultiple[i] != "0" && txtPOQtyMultiple[i] != "undefined" ? Int32.Parse(txtPOQtyMultiple[i]) : tempdetail.Poqty;
                        tempdetail.Eta = txtETAMultiple[i]!= null && txtETAMultiple[i] != "undefined" ? DateTime.Parse(txtETAMultiple[i]) : tempdetail.Eta;
                        tempdetail.Etd = txtETDMultiple[i]!= null && txtETDMultiple[i] != "undefined" ? DateTime.Parse(txtETDMultiple[i]) : tempdetail.Etd;
                    
                        await _context.SaveChangesAsync();

                        temppassdata.Eta = date;
                        temppassdata.Poqty = qty;
                        temppassdata.Etd = etd;
                        temppassdata.ExpectedPoissuedDate = podate;
                    }

                    data.Add(tempdetail);
                    temp.Add(temppassdata);

                }
                for (var i = 0; i < idnew.Length; i++)
                {
                    temp[i].Eta = temp[i].Eta != data[i].Eta ? data[i].Eta : null;
                    temp[i].Poqty = temp[i].Poqty != data[i].Poqty ? data[i].Poqty : null;
                    temp[i].ExpectedPoissuedDate = temp[i].ExpectedPoissuedDate != data[i].ExpectedPoissuedDate ? data[i].ExpectedPoissuedDate : null;
                    temp[i].Etd = temp[i].Etd != data[i].Etd ? data[i].Etd : null;

                    passdata.Add(data[i]);
                    passdata[i].Eta = temp[i].Eta;
                    passdata[i].Poqty = temp[i].Poqty;
                    passdata[i].ExpectedPoissuedDate = temp[i].ExpectedPoissuedDate;
                    passdata[i].Etd = temp[i].Etd;
                }
            }

            await SendNotifByType("", data[0].PreProcId, "UpdatedPODetailFromProc", passdata, new List<PreProcDetail>(), new PreProcRiskAssesmentDatum(), "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcPodatum>> PostUpdateGRData(int procid, string[] idnew, string[] poline, string inputtype, string txtGRDateSingle, string txtGRQtySingle, string[] txtGRDateMultiple, string[] txtGRQtyMultiple, IFormFile postedfilegrn)
        {
            getUserNow();
            var data = new List<PreProcPodatum>();
            var passdata = new List<PreProcPodatum>();
            var temp = new List<PreProcPodatum>();

            if (inputtype == "Single")
            {
                for (var i = 0; i < idnew.Length; i++)
                {
                    var tempdetail = await _context.PreProcPodata.FirstOrDefaultAsync(x => x.PreProcId == procid && x.PolineId == poline[i] && x.Idnew == idnew[i]);
                    var temppassdata = new PreProcPodatum();
                    var qty = tempdetail.Grqty;
                    var date = tempdetail.Grdate;

                    if (tempdetail != null)
                    {
                        tempdetail.Grdate = txtGRDateSingle != null ? DateTime.Parse(txtGRDateSingle) : tempdetail.Grdate;
                        tempdetail.Grqty = txtGRQtySingle != null && txtGRQtySingle != "0" ? Int32.Parse(txtGRQtySingle) : tempdetail.Grqty;

                        await _context.SaveChangesAsync();

                        temppassdata.Grdate = date;
                        temppassdata.Grqty = qty;
                    }
                    data.Add(tempdetail);
                    temp.Add(temppassdata);
                }
                for (var i = 0; i < idnew.Length; i++)
                {
                    temp[i].Grdate = temp[i].Grdate != data[i].Grdate ? data[i].Grdate : null;
                    temp[i].Grqty = temp[i].Grqty != data[i].Grqty ? data[i].Grqty : null;

                    passdata.Add(data[i]);
                    passdata[i].Grdate = temp[i].Grdate;
                    passdata[i].Grqty = temp[i].Grqty;
                }
            }
            else
            {
                for (var i = 0; i < idnew.Length; i++)
                {
                    var tempdetail = await _context.PreProcPodata.FirstOrDefaultAsync(x => x.PreProcId == procid && x.PolineId == poline[i] && x.Idnew == idnew[i]);
                    var qty = tempdetail.Grqty;
                    var date = tempdetail.Grdate;
                    var temppassdata = new PreProcPodatum();

                    if (tempdetail != null)
                    {
                        tempdetail.Grdate = txtGRDateMultiple[i] != null ? DateTime.Parse(txtGRDateMultiple[i]) : tempdetail.Grdate;
                        tempdetail.Grqty = txtGRQtyMultiple[i] != null && txtGRQtyMultiple[i] != "0" ? Int32.Parse(txtGRQtyMultiple[i]) : tempdetail.Grqty;

                        await _context.SaveChangesAsync();
                        
                        temppassdata.Grdate = date;
                        temppassdata.Grqty = qty;
                    }
                    data.Add(tempdetail);
                    temp.Add(temppassdata);
                }

                for (var i = 0; i < idnew.Length; i++)
                {
                    temp[i].Grdate = temp[i].Grdate != data[i].Grdate ? data[i].Grdate : null;
                    temp[i].Grqty = temp[i].Grqty != data[i].Grqty ? data[i].Grqty : null;

                    passdata.Add(data[i]);
                    passdata[i].Grdate = temp[i].Grdate;
                    passdata[i].Grqty = temp[i].Grqty;
                }
            }
            await SendNotifByType("", passdata[0].PreProcId, "UpdatedFromWH", passdata, new List<PreProcDetail>(), new PreProcRiskAssesmentDatum(), "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());

            return data;
        }

        [HttpPost]
        public async Task<PreProcRiskAssesmentDatum> PostAssignRiskCommittee(int procid, string resultSelectedCommitteeDomain)
        {
            getUserNow();
            var data = new PreProcRiskAssesmentDatum();

            var riskassess = await _context.PreProcRiskAssesmentData.FirstOrDefaultAsync(x => x.PreProcId == procid);
            if (riskassess != null)
            {
                riskassess.EmailAssignTo = resultSelectedCommitteeDomain;
                await _context.SaveChangesAsync();
            }
            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> PostAssignPM(int procid, string[] pm, int[] detailid)
        {
            getUserNow();
            var data = new List<PreProcDetail>();

            for (var i = 0; i < pm.Length; i++)
            {
                var detail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid[i] && x.Id == procid);

                if (detail != null)
                {
                    detail.ProductManager = pm[i];
                    await _context.SaveChangesAsync();

                    data.Add(detail);
                }
            }
            await SendNotifByType("", procid, "AssignProdMgr", new List<PreProcPodatum>(), new List<PreProcDetail>(), new PreProcRiskAssesmentDatum(), "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());

            return data;
        }

        [HttpPost]
        public async Task<MsParameterValue> GetValueFromTableParameter(string title, string parameter)
        {
            getUserNow();
            var data = new MsParameterValue();

            var paramvalue = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == title && x.Parameter == parameter);
            if (paramvalue != null)
            {
                data = paramvalue;
            }
            return data;
        }

        [HttpPost]
        public async Task<List<CogsrevNpc>> GetCSDataFromSearch(string pid)
        {
            getUserNow();
            var data = new List<CogsrevNpc>();
            var paramvalue = new List<CogsrevNpc>();

            if (pid != null)
            {
                if (pid.Length == 10)
                {
                    paramvalue = await _context.CogsrevNpcs.Where(x => x.ProjectId.ToLower() == pid).ToListAsync();
                }
                else if (pid.Length == 13)
                {
                    paramvalue = await _context.CogsrevNpcs.Where(x => x.Pidfull.ToLower() == pid).ToListAsync();
                }
            }
            

            if (paramvalue != null)
            {
                data = paramvalue;
            }
            return data;
        }

        [HttpPost]
        public async Task<List<PreProcPodatum>> CheckExistingPOData(string poline)
        {
            getUserNow();
            var data = new List<PreProcPodatum>();
            var paramvalue = new List<PreProcPodatum>();

            if (poline != null && poline != "")
            {
                paramvalue = await _context.PreProcPodata.Where(x => x.PolineId.ToLower() == poline.ToLower()).ToListAsync();
            }


            if (paramvalue != null)
            {
                data = paramvalue;
            }
            return data;
        }

        [HttpPost]
        public async Task<List<PreProcDetail>> CheckExistingCSData(string pidfull)
        {
            getUserNow();
            var data = new List<PreProcDetail>();
            var paramvalue = new List<PreProcDetail>();

            if (pidfull != null && pidfull != "")
            {
                paramvalue = await _context.PreProcDetails.Where(x => x.Pidfull.ToLower() == pidfull.ToLower()).ToListAsync();
            }


            if (paramvalue != null)
            {
                data = paramvalue;
            }
            return data;
        }

        [HttpPost]
        public async Task<PreProcApprovalHistory> GetDetailApprovalItem(int idapproval, string role)
        {
            getUserNow();

            var approval = await _context.PreProcApprovalHistories.OrderByDescending(x=>x.Id).FirstOrDefaultAsync(x => x.DetailId == idapproval && x.ApproverRole.Contains(role));

            var data = approval != null ? approval : new PreProcApprovalHistory();

            return data;
        }

        [HttpPost]
        public async Task<List<VwGetMplPoAll>> GetItemPOOracle(string pid)
        {
            getUserNow();
            var data = new List<VwGetMplPoAll>();
            var paramvalue = new List<VwGetMplPoAll>();

            if (pid != null)
            {
                var splitpid = pid.Split(";");
                var tempsplitpid = splitpid.Distinct().ToList();

                foreach (var item in tempsplitpid)
                {
                    if (item != "" && item != null)
                    {
                        var test = await _context.VwGetMplPoAlls.Where(x => x.Pid.ToLower() == item.ToLower()).ToListAsync();
                        paramvalue.AddRange(test);
                    }

                }
                //try
                //{
                //    ViewBag.param = from t in _context.VwGetMplPoAlls.ToList()
                //                                  select t.NoPo;
                //    //await _context.VwGetMplPoAlls.Select(s => new { name = s.NoPo, index = s.NamaVendor }).ToListAsync();
                //}
                //catch (Exception exec)
                //{
                //    Console.WriteLine(exec.Message);
                //    throw;
                //}
            }

            if (paramvalue != null)
            {
                data = paramvalue;
            }
            return data;
        }

        [HttpPost]
        public async Task<List<PreProcPodatum>> GetItemPreProcPO(string procid, string idnew, string[] polineid)
        {
            getUserNow();
            var data = new List<PreProcPodatum>();

            var tes = await _context.PreProcPodata.Where(x => x.PreProcId == Int32.Parse(procid)).ToListAsync();

            if (polineid.Length == 0)
            {
                tes = await _context.PreProcPodata.Where(x => x.PreProcId == Int32.Parse(procid) && x.Idnew == idnew).ToListAsync();
                data.AddRange(tes);
            }
            else
            {
                foreach (var item in polineid)
                {
                    var tempsplit = item.Split(";");

                    if (tempsplit.Length > 0)
                    {
                        tes = await _context.PreProcPodata.Where(x => x.PreProcId == Int32.Parse(procid) && x.PolineId == tempsplit[0] && x.Idnew == tempsplit[1]).ToListAsync();
                    }

                    data.AddRange(tes);
                }
                
            }

            return data;
        }

        [HttpPost]
        public async Task<List<PreProcFileAttachment>> ShowFileAttachment(string procid, string doctype, int detailid)
        {
            getUserNow();
            var datas = new List<PreProcFileAttachment>();

            //var doctypes = from t in _context.MsDocTypes
            //               select t;

            //var tempdoctype = await _context.MsDocTypes.FirstOrDefaultAsync(t => t.DocType == doctype);

            if (doctype == "GRN")
            {
                var tes = await _context.PreProcFileAttachments.Where(x => x.PreProcId == Int32.Parse(procid) && x.Bastid == detailid && x.DocumentTypeId == 19).ToListAsync();

                datas.AddRange(tes);
            }
            return datas;
        }

        [HttpPost]
        public async Task<List<PreProcPodatum>> GetItemPreProcPOforBrowse(string procid)
        {
            getUserNow();
            var data = new List<PreProcPodatum>();

            var tes = await _context.PreProcPodata.Where(x => x.PreProcId == Int32.Parse(procid) && x.PolineId != null).ToListAsync();

            data = tes;

            return data;
        }

        [HttpPost]
        public List<PreProcDetail> ImportExcel<T>(string excelFilePath, string sheetName)
        {
            getUserNow();
            //List<T> list = new List<T>();
            List<PreProcDetail> list2 = new List<PreProcDetail>();
            PreProcDetail model;

            Type typeOfObject = typeof(T);
            using (IXLWorkbook workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheets.Where(w => w.Name == sheetName).First();
                var properties = typeOfObject.GetProperties();
                //header column text
                var columns = worksheet.FirstRow().Cells().Select((v, i) => new { Value = v.Value, Index = i + 1 }); //index start from 1
                int emptyrow = 0;
                foreach (IXLRow row in worksheet.RowsUsed().Skip(1))
                {
                    //T obj = (T)Activator.CreateInstance(typeOfObject);
                    model = new PreProcDetail();

                    string tes = row.Cell(1).Value.ToString();
                    if (tes != "")
                    {
                        foreach (var prop in properties)
                        {
                            string b = prop.Name.ToString().ToLower();
                            int colIndex = columns.SingleOrDefault(c => c.Value.ToString().ToLower() == b).Index;
                            string val = "";
                            if (colIndex == 3)
                            {
                                val = row.Cell(colIndex).CachedValue.ToString().Replace("Text ", "");
                            }
                            else
                            {
                                val = row.Cell(colIndex).Value.ToString();
                            }
                            var type = prop.PropertyType;
                            switch (colIndex)
                            {
                                case 1:
                                    model.Id = Int32.Parse(val);
                                    break;
                                case 2:
                                    model.Principle = val;
                                    break;
                                case 3:
                                    model.Nkpcode = val;
                                    break;
                                case 4:
                                    model.ItemDesc = val;
                                    break;
                                case 5:
                                    model.Qty = double.Parse(val);
                                    break;
                                case 6:
                                    model.Warranty = Int32.Parse(val);
                                    break;
                                case 7:
                                    model.Insurance = val == "Y" ? true : false;
                                    break;
                                case 8:
                                    model.SurDukCheck = val == "Y" ? true : false;
                                    break;
                                default:
                                    break;
                            }

                            //prop.SetValue(obj, val.ToString());
                        }
                    list2.Add(model);
                    }
                    else
                    {
                        if (emptyrow > 0)
                        {
                            break;
                        }
                        emptyrow++;
                    }
                }
            }

            return list2;
        }

        
        public async Task<IActionResult> NewItem()
        {
            getUserNow();

            var useremail = getUserNowEmail();

            var paramemail = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Current Login Email").FirstOrDefaultAsync();

            if (paramemail.AlphaNumericValue != null && paramemail.AlphaNumericValue != "")
            {
                useremail = paramemail.AlphaNumericValue;
            }

            //useremail = "hananto.wibowo@multipolar.com";      //PM MEMBER
            //useremail = "kevin.christian@multipolar.com";     //AM MEMBER
            //useremail = "ferianto.hanafie@multipolar.com";    //PROC MEMBER
            //useremail = "dini.hayati@multipolar.com";           //PS MEMBER
            //useremail = "solehudin@multipolar.com";           //WH MEMBER
            //useremail = "arief@multipolar.com";           //TEST COMMITTEE MEMBER
            //useremail = "IVAN@multipolar.com";           //BOD MEMBER

            var useraccessAM = await _context.VwGetAmmembers.FirstOrDefaultAsync(x=>x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
            ViewBag.useraccessAM = useraccessAM;
            var useraccessTA = await _context.VwGetTenderAdminMemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
            ViewBag.useraccessTA = useraccessTA;
            var useraccessPMO = await _context.VwGetPmomemberFulls.FirstOrDefaultAsync(x => x.EmployeeEmail.ToLower().Equals(useremail.ToLower()));
            var useraccessSuperAdmin = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title.Equals("PreProc") && x.Parameter.Equals("Super Admin Access") && x.AlphaNumericValue.ToLower().Equals(useremail.ToLower()));
            

            if (useraccessAM != null || useraccessTA != null || useraccessSuperAdmin != null || useraccessPMO != null)
            {
                PreProcContextProcedures dbsp = new PreProcContextProcedures(_context);
                //var tes = dbsp.sp_GetCustomerAsync().Result.ToList();
                ViewBag.amemployees = await _context.VwGetAmmembers.OrderBy(x=>x.EmployeeName).ToListAsync();
                ViewBag.technicalemployees = await _context.VwGetPresalesMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
                ViewBag.tenderemployees = await _context.VwGetTenderAdminMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
                ViewBag.customers = await _context.MsCustomerVws.OrderBy(x => x.AccountName).ToListAsync();
                ViewBag.picprocs = await _context.VwGetProcurementMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
                ViewBag.projecttypes = await _context.MsProjectTypes.ToListAsync();
                ViewBag.presales = await _context.VwGetAllOptyIds.OrderByDescending(x=>x.Id).ToListAsync();
                ViewBag.pids = await _context.VwGetAllPids.OrderByDescending(x=>x.TrProjectIdId).ToListAsync();
                ViewBag.items = await _context.MsPreProcStages.ToListAsync();
                ViewBag.projectsupp = await _context.VwGetProjectSupportMemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
                ViewBag.pm = await _context.VwGetPmmemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
                ViewBag.pc = await _context.VwGetPcmemberFulls.OrderBy(x => x.EmployeeName).ToListAsync();
                ViewBag.bod = await _context.VwGetBodfulls.OrderBy(x => x.Nama).ToListAsync();
                ViewBag.defaultbod = await _context.VwGetBodfulls.Where(x => x.CurrentGroup.Equals("EA")).OrderBy(x => x.Nama).ToListAsync();
                ViewBag.urlgenpid = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title.Equals("PreProc") && x.Parameter.Equals("Url Generate PID"));

                return View();
            }
            else
            {
                return RedirectToAction(nameof(Index));
            }

        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> PostNewItem(PreProcGeneralInfo2 model, IFormFile postedfilerfp, IFormFile postedfileboq, IFormFile postedfilequot, IFormFile postedfilegp, IFormFile postedfileproposal, IFormFile postedfileothers, bool NeedWarranty, bool Hedging)
        {
            getUserNow();

            PreProcGeneralInfo2 preProcGI = new PreProcGeneralInfo2();//= new preProcGI();
            PreProcHeaderAttachment preProcAttach;//= new preProcAttach();
            var existproc = await _context.PreProcGeneralInfo2s.OrderByDescending(x => x.Id).FirstOrDefaultAsync();
            //var bod = model.EmailCommitteeDefault.Split(";");
            //if (ModelState.IsValid)
            //{
            if (model.Stage == "RiskAssessment")
            {
                preProcGI.PresalesId = model.PresalesId;
                preProcGI.Pid = model.PresalesId;

                preProcGI.Type = "Waiting Risk Assessment";
            }
            else if (model.Stage == "Presales")
            {
                preProcGI.PresalesId = model.PresalesId;
                preProcGI.Pid = model.PresalesId;
                preProcGI.Type = "Normal";
                preProcGI.Pc = model.Pc;
                preProcGI.Pm = model.Pm;
                preProcGI.ProjectSupport = model.ProjectSupport;
            }
            else
            {
                preProcGI.PresalesId = model.Pid;
                preProcGI.Pid = model.Pid;
                preProcGI.Type = "Normal";
                preProcGI.Pc = model.Pc;
                preProcGI.Pm = model.Pm;
                preProcGI.ProjectSupport = model.ProjectSupport;
            }

            preProcGI.Stage = model.Stage;
            preProcGI.RfpId = model.RfpId;
            preProcGI.ProposalId = model.ProposalId;
            preProcGI.QuotationId = model.QuotationId;
            preProcGI.Customer = model.Customer;
            preProcGI.Am = model.Am;
            preProcGI.PropMgr = model.PropMgr;
            preProcGI.TenderAdm = model.TenderAdm;
            preProcGI.ProposalSubmissionDate = model.ProposalSubmissionDate;
            preProcGI.QuotationDeadline = model.QuotationDeadline;
            preProcGI.Hedging = Hedging;
            preProcGI.CreatedBy = User.Identity.Name;
            preProcGI.CreationDate = DateTime.Now;
            preProcGI.Created = DateTime.Now;
            preProcGI.NeedWarranty = NeedWarranty;

            //decimal resultidr = 0;
            //decimal resultusd = 0;
            //decimal resultgp = 0;

            //if (model.ProjectValueIdr.HasValue)
            //{
            //    string decimalString = model.ProjectValueIdr.Value.ToString().Replace(",", ".");
            //    resultidr = decimal.Parse(decimalString, CultureInfo.InvariantCulture);
            //}
            //if (model.ProjectValueUsd.HasValue)
            //{
            //   string decimalString = model.ProjectValueUsd.Value.ToString().Replace(",", ".");
            //    resultusd = decimal.Parse(decimalString, CultureInfo.InvariantCulture);
            //}
            //if (model.EstGpgeneral.HasValue)
            //{
            //    string decimalString = model.EstGpgeneral.Value.ToString().Replace(",", ".");
            //    resultgp = decimal.Parse(decimalString, CultureInfo.InvariantCulture);
            //}

            preProcGI.ProjectValueIdr = model.ProjectValueIdr;
            preProcGI.ProjectValueUsd = model.ProjectValueUsd;
            preProcGI.EstGpgeneral = model.EstGpgeneral;
            preProcGI.ProjectType = model.ProjectType;
            preProcGI.Description = model.Description;
            preProcGI.PicProc = model.PicProc;

            var bod = "";
            if (model.EmailCommitteeDefault == "" || model.EmailCommitteeDefault == null)
            {
                var currgroup = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Parameter == "Default Directorate BOD");
                var tempbod = await _context.VwGetBodfulls.FirstOrDefaultAsync(x=>x.CurrentGroup == currgroup.AlphaNumericValue);
                if (tempbod != null)
                {
                    bod = tempbod.EmailAddress;
                }
            }
            else
            {
                bod = model.EmailCommitteeDefault;
            }
            preProcGI.EmailCommitteeDefault = bod;
            preProcGI.ListItemId = existproc.ListItemId + 1;

            _context.Add(preProcGI);
            await _context.SaveChangesAsync();

            if (postedfilerfp != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                                //+
                                                                postedfilerfp.FileName);

                var doctype = from t in _context.MsDocTypes
                                select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("RFP")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilerfp.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilerfp.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfileboq != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                                //+
                                                                postedfileboq.FileName);

                var doctype = from t in _context.MsDocTypes
                                select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("BOQ")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfileboq.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfileboq.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilequot != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                                //+
                                                                postedfilequot.FileName);

                var doctype = from t in _context.MsDocTypes
                                select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("QUO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilequot.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilequot.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilegp != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                                //+
                                                                postedfilegp.FileName);

                var doctype = from t in _context.MsDocTypes
                                select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("GPJ")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilegp.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilegp.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfileproposal != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                                //+
                                                                postedfileproposal.FileName);

                var doctype = from t in _context.MsDocTypes
                                select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("PRO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfileproposal.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfileproposal.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfileothers != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                                //+
                                                                postedfileothers.FileName);

                var doctype = from t in _context.MsDocTypes
                                select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("OTH")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfileothers.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfileothers.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }//} 

            await SendNotifByType("", preProcGI.Id, "NewItem", new List<PreProcPodatum>(), new List<PreProcDetail>(), new PreProcRiskAssesmentDatum(), "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());
                
            return RedirectToAction("EditDetail", new { id = preProcGI.Id });

        }

        [HttpPost]
        public async Task<PreProcPendingBillingReason> SendNotifByType(string message, int procid, string type, List<PreProcPodatum> list, List<PreProcDetail> listtempproc, PreProcRiskAssesmentDatum riskdata, string overallrisk, List<PreProcRiskDatum> listriskdata, List<PreProcPrincipalProjectRevenue> listprojectrev)
        {
            getUserNow();
            var data = new PreProcPendingBillingReason();

            string TextOpen = @"<html>
                                      <head>
                                         <style>
                                        <!--
                                         /* Font Definitions */
                                         @font-face
	                                        {font-family:""Cambria Math"";
	                                        panose-1:2 4 5 3 5 4 6 3 2 4;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:roman;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:3 0 0 0 1 0;}
                                        @font-face
	                                        {font-family:Calibri;
	                                        panose-1:2 15 5 2 2 2 4 3 2 4;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:swiss;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:-469750017 -1073732485 9 0 511 0;}
                                        @font-face
	                                        {font-family:""Segoe UI"";
	                                        panose-1:2 11 5 2 4 2 4 2 2 3;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:swiss;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:-469750017 -1073683329 9 0 511 0;}
                                         /* Style Definitions */
                                         p.MsoNormal, li.MsoNormal, div.MsoNormal
	                                        {mso-style-unhide:no;
	                                        mso-style-qformat:yes;
	                                        mso-style-parent:"""";
	                                        margin:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        a:link, span.MsoHyperlink
	                                        {mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        color:blue;
	                                        text-decoration:underline;
	                                        text-underline:single;}
                                        a:visited, span.MsoHyperlinkFollowed
	                                        {mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        color:purple;
	                                        text-decoration:underline;
	                                        text-underline:single;}
                                        p
	                                        {mso-style-priority:99;
	                                        mso-margin-top-alt:auto;
	                                        margin-right:0in;
	                                        mso-margin-bottom-alt:auto;
	                                        margin-left:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        p.msonormal0, li.msonormal0, div.msonormal0
	                                        {mso-style-name:msonormal;
	                                        mso-style-unhide:no;
	                                        mso-margin-top-alt:auto;
	                                        margin-right:0in;
	                                        mso-margin-bottom-alt:auto;
	                                        margin-left:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        .MsoChpDefault
	                                        {mso-style-type:export-only;
	                                        mso-default-props:yes;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-ascii-font-family:Calibri;
	                                        mso-ascii-theme-font:minor-latin;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;
	                                        mso-hansi-font-family:Calibri;
	                                        mso-hansi-theme-font:minor-latin;
	                                        mso-bidi-font-family:""Times New Roman"";
	                                        mso-bidi-theme-font:minor-bidi;}
                                        @page WordSection1
	                                        {size:8.5in 11.0in;
	                                        margin:1.0in 1.0in 1.0in 1.0in;
	                                        mso-header-margin:.5in;
	                                        mso-footer-margin:.5in;
	                                        mso-paper-source:0;}
                                        div.WordSection1
	                                        {page:WordSection1;}
                                         table.MsoNormalTable
	                                        {mso-style-name:""Table Normal"";
	                                        mso-tstyle-rowband-size:0;
	                                        mso-tstyle-colband-size:0;
	                                        mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        mso-style-parent:"""";
	                                        mso-padding-alt:0in 5.4pt 0in 5.4pt;
	                                        mso-para-margin:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-ascii-font-family:Calibri;
	                                        mso-ascii-theme-font:minor-latin;
	                                        mso-hansi-font-family:Calibri;
	                                        mso-hansi-theme-font:minor-latin;
	                                        mso-bidi-font-family:""Times New Roman"";
	                                        mso-bidi-theme-font:minor-bidi;}
                                        </style>
                                       </head>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>Dear ";

            string TextOpenAfterHead = @",</span>
                                        </p>";

            string TextClosedWithTable = @"</table><br>";

            const string TextClosed = @"<br><br><p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>&nbsp; <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US'>
                                            <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>Please Do Not Reply to this email because we are not monitoring this inbox<o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <br>
                                        <div class='elementToProof'>
                                          <div style='font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);'>
                                            <br>
                                          </div>
                                          <div id='Signature'>
                                            <div>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <span style='font-size:8pt;font-family:Times New Roman,serif'>Best Regards,&nbsp;</span>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='font-family:Times New Roman,serif'>Internal System&nbsp;</span>
                                                </b>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'>PT MULTIPOLAR TECHNO</span>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> LOGY Tbk(MLPT)</span>
                                                </b>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> &nbsp;</span>
                                                </b>
                                              </p>
                                            </div>
                                          </div>
                                        </div>
                                        </body>
                                    </html>";

            string TdOpener = @"<td style='border:none;padding:3.0pt 3.0pt 3.0pt 3.0pt'><p class=MsoNormal><span style='font-family:""Segoe UI"",sans-serif'>";
            string TdCloser = @"<o:p></o:p></span></p></td>";
            //------------------------------------------------------------------------------------------------------------------------ REMINDER DELIVERY ORDER BORROW ---------------------------------------------------------------------------------------
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "10.10.62.17";
                builder.UserID = "sa";
                builder.Password = "Password1!";
                builder.InitialCatalog = "eRequisition";
                builder.TrustServerCertificate = true;
                builder.MultipleActiveResultSets = true;
                var proc = _context.PreProcGeneralInfo2s.FirstOrDefault(x => x.Id == procid);
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    var mailMessage = new MailMessage();
                    mailMessage.Bcc.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("adam.takariyanto@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("internal.apps@multipolar.com"));
                    mailMessage.From = new MailAddress("mlpt365@multipolar.com", "Multipolar Technology 365");
                    mailMessage.Subject = "[PREPROC] Assignment Member Project ID " + proc.Pid;

                    if (type == "RemindProcWH" || type == "PurchaseReq")
                    {
                        mailMessage.Subject = "[PREPROC] Reminder Project ID " + proc.Pid;
                    }
                    else if (type == "UpdatedFromWH" || type == "UpdatedPODetailFromProc" || type == "UpdatedPOFromProc")
                    {
                        mailMessage.Subject = "[PREPROC] Updated Data Project ID " + proc.Pid;
                    }
                    else if (type == "UpdatedRiskAssess")
                    {
                        mailMessage.Subject = "[PREPROC] Updated Risk Assessment Data Project ID " + proc.Pid;
                    }

                    mailMessage.Body = TextOpen;
                    mailMessage.IsBodyHtml = true;

                    string TableHeaderRemark = @"<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=0 style='mso-cellspacing:1.5pt;border:solid #104E8B 1.0pt;mso-border-alt:solid #104E8B .75pt; mso-yfti-tbllook:1184;mso-padding-alt:3.0pt 3.0pt 3.0pt 3.0pt'>
                                        <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>PreProcID<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>No<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Principal/Manufacturer<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Item Description<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Item Remark<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Remarks<o:p></o:p></span></b></p>
                                            </td></tr>";

                    string TableHeaderPO = @"<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=0 style='mso-cellspacing:1.5pt;border:solid #104E8B 1.0pt;mso-border-alt:solid #104E8B .75pt; mso-yfti-tbllook:1184;mso-padding-alt:3.0pt 3.0pt 3.0pt 3.0pt'>
                                        <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>PreProcID<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>No<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>PO Number<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>PO Item<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>NKS<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>PO Qty<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>ETA<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Expected PO Issued Date<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>ETD<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>GR Date<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>GR Qty<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Remarks<o:p></o:p></span></b></p>
                                            </td></tr>";

                    string TableHeaderProjRev = @"<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=0 style='mso-cellspacing:1.5pt;border:solid #104E8B 1.0pt;mso-border-alt:solid #104E8B .75pt; mso-yfti-tbllook:1184;mso-padding-alt:3.0pt 3.0pt 3.0pt 3.0pt'>
                                        <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>SubCon Name<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>%<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>High Level Scop<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Preload<o:p></o:p></span></b></p>
                                            </td></tr>";
                    string CurrentGroup = "";
                    string Status = "";
                    string Level2 = "";
                    string Level4 = "";
                    string Level5 = "";
                    int Level = 0;

                    //var splitemail = proc.PicProc.Replace(" ", "").Split(";");
                    var splitemail = proc.PicProc.Replace(" ", "").Split(";");

                    var additionrecpt = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT");
                    var additionrecptpmo = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT PMO").ToListAsync();

                    foreach (var item in splitemail)
                    {
                        if (item != null && item != "")
                        {
                            String sql = "SELECT * FROM gp_ms_Emp WHERE Login = '" + item + "'";
                            using (SqlCommand command = new SqlCommand(sql, connection))
                            {
                                using (SqlDataReader reader = command.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        if (splitemail.Length == 1)
                                        {
                                            mailMessage.Body += reader[1].ToString();
                                        }
                                        //mailMessage.To.Add(new MailAddress(reader[18].ToString()));
                                        Status = reader[12].ToString();
                                        Level2 = reader[4].ToString();
                                        Level4 = reader[7].ToString();
                                        Level5 = reader[9].ToString();
                                        Level = Convert.ToInt32(reader[13].ToString());
                                        if (Status == "m")
                                        {
                                            if (Level == 6)
                                            {
                                                CurrentGroup = Level5;
                                            }
                                            else if (Level == 5)
                                            {
                                                CurrentGroup = Level4;
                                            }
                                            else if (Level == 4)
                                            {
                                                CurrentGroup = Level2;
                                            }
                                        }
                                        else
                                        {
                                            CurrentGroup = reader[14].ToString();
                                        }
                                        mailMessage.To.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));

                                        if (additionrecpt != null)
                                        {
                                            mailMessage.To.Add(new MailAddress(additionrecpt.AlphaNumericValue));

                                        }
                                        


                                    }
                                }
                            }
                        }
                        
                    }
                    if (splitemail.Length > 1)
                    {
                        mailMessage.Body += "All";
                    }

                    if (type == "RemindProcWH")
                    {
                        mailMessage.Body += TextOpenAfterHead + "<br>" + message + "<br><br> ";
                    }
                    else if (type == "NewItem")
                    {
                        mailMessage.Body += TextOpenAfterHead + "<br>" + message + "<br><br> You have been assigned to this project. <br> ";
                    }
                    else if (type == "AssignProdMgr")
                    {
                        mailMessage.Body += TextOpenAfterHead + "<br>" + "You have been assigned to this project. <br> ";
                    }
                    else if (type == "PurchaseReq")
                    {
                        mailMessage.Body += TextOpenAfterHead + "<br>" + message + "<br><br>";
                    }
                    else if (type == "NotifToAM")
                    {
                        mailMessage.Body += TextOpenAfterHead + "<br>" + "PMO Checklist has been updated. <br>";
                    }
                    else if (type == "UpdatedFromWH")
                    {
                        if (additionrecptpmo != null)
                        {
                            foreach (var addrecptpmo in additionrecptpmo)
                            {
                                mailMessage.To.Add(new MailAddress(addrecptpmo.AlphaNumericValue));
                            }
                        }
                        mailMessage.Body += TextOpenAfterHead + "<br>List Updated Data <br>" + TableHeaderPO;
                        if (list.Count > 0)
                        {
                            for (var i = 0; i < list.Count; i++)
                            {
                                if (list[i].Grdate != null || list[i].Grqty != null)
                                {
                                    var notedate = list[i].Grdate != null ? "-Updated GR Date <br>" : "";
                                    var noteqty = list[i].Grqty != null ? "-Updated GR Qty" : "";

                                    mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" + 
                                                        TdOpener + list[i].PreProcId + TdCloser +
                                                        TdOpener + list[i].Idnew + TdCloser +
                                                        TdOpener + list[i].Poid + TdCloser +
                                                        TdOpener + list[i].ItemDescription + TdCloser +
                                                        TdOpener + list[i].ItemNo + TdCloser +
                                                        TdOpener + list[i].Poqty + TdCloser +
                                                        TdOpener + list[i].Eta + TdCloser +
                                                        TdOpener + list[i].ExpectedPoissuedDate + TdCloser +
                                                        TdOpener + list[i].Etd + TdCloser +
                                                        TdOpener + list[i].Grdate + TdCloser +
                                                        TdOpener + list[i].Grqty + TdCloser +
                                                        TdOpener + notedate + noteqty + TdCloser + "</tr>";
                                }
                                
                            }
                        }
                        
                    }
                    else if (type == "UpdatedRemarkFromProc")
                    {
                        if (additionrecptpmo != null)
                        {
                            foreach (var addrecptpmo in additionrecptpmo)
                            {
                                mailMessage.To.Add(new MailAddress(addrecptpmo.AlphaNumericValue));
                            }
                        }

                        mailMessage.Body += TextOpenAfterHead + "<br>List Updated Data <br>" + TableHeaderRemark;
                        if (listtempproc.Count > 0)
                        {
                            for (var i = 0; i < listtempproc.Count; i++)
                            {
                                if (listtempproc[i].RemarkOnItem != null )
                                {
                                    var noteremark = listtempproc[i].RemarkOnItem != null ? "-Updated Remark Item" : "";

                                    mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" +
                                                        TdOpener + listtempproc[i].Id + TdCloser +
                                                        TdOpener + listtempproc[i].Idnew + TdCloser +
                                                        TdOpener + listtempproc[i].Principle + TdCloser +
                                                        TdOpener + listtempproc[i].ItemDesc + TdCloser +
                                                        TdOpener + listtempproc[i].RemarkOnItem + TdCloser +
                                                        TdOpener + noteremark + TdCloser + "</tr>";
                                }

                            }
                        }
                    }
                    else if (type == "UpdatedPOFromProc")
                    {
                        if (additionrecptpmo != null)
                        {
                            foreach (var addrecptpmo in additionrecptpmo)
                            {
                                mailMessage.To.Add(new MailAddress(addrecptpmo.AlphaNumericValue));
                            }
                        }

                        mailMessage.Body += TextOpenAfterHead + "<br>List Updated Data <br>" + TableHeaderPO;
                        if (list.Count > 0)
                        {
                            for (var i = 0; i < list.Count; i++)
                            {
                                if (list[i].Poid != null)
                                {
                                    var notepo = list[i].Poid != null ? "-Added PO Number" : "";

                                    mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" +
                                                        TdOpener + list[i].PreProcId + TdCloser +
                                                        TdOpener + list[i].Idnew + TdCloser +
                                                        TdOpener + list[i].Poid + TdCloser +
                                                        TdOpener + list[i].ItemDescription + TdCloser +
                                                        TdOpener + list[i].ItemNo + TdCloser +
                                                        TdOpener + list[i].Poqty + TdCloser +
                                                        TdOpener + list[i].Eta + TdCloser +
                                                        TdOpener + list[i].ExpectedPoissuedDate + TdCloser +
                                                        TdOpener + list[i].Etd + TdCloser +
                                                        TdOpener + list[i].Grdate + TdCloser +
                                                        TdOpener + list[i].Grqty + TdCloser +
                                                        TdOpener + notepo + TdCloser + "</tr>";
                                }

                            }
                        }
                    }
                    else if (type == "UpdatedPODetailFromProc")
                    {
                        if (additionrecptpmo != null)
                        {
                            foreach (var addrecptpmo in additionrecptpmo)
                            {
                                mailMessage.To.Add(new MailAddress(addrecptpmo.AlphaNumericValue));
                            }
                        }

                        mailMessage.Body += TextOpenAfterHead + "<br>List Updated Data <br>" + TableHeaderPO;
                        if (list.Count > 0)
                        {
                            for (var i = 0; i < list.Count; i++)
                            {
                                if ((list[i].Eta != null || list[i].Poqty != null) && (list[i].Etd == null || list[i].ExpectedPoissuedDate == null))
                                {
                                    var notedate = list[i].Eta != null ? "-Updated Eta <br>" : "";
                                    var noteqty = list[i].Poqty != null ? "-Updated PO Qty" : "";

                                    mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" +
                                                        TdOpener + list[i].PreProcId + TdCloser +
                                                        TdOpener + list[i].Idnew + TdCloser +
                                                        TdOpener + list[i].Poid + TdCloser +
                                                        TdOpener + list[i].ItemDescription + TdCloser +
                                                        TdOpener + list[i].ItemNo + TdCloser +
                                                        TdOpener + list[i].Poqty + TdCloser +
                                                        TdOpener + list[i].Eta + TdCloser +
                                                        TdOpener + list[i].ExpectedPoissuedDate + TdCloser +
                                                        TdOpener + list[i].Etd + TdCloser +
                                                        TdOpener + list[i].Grdate + TdCloser +
                                                        TdOpener + list[i].Grqty + TdCloser +
                                                        TdOpener + notedate + noteqty + TdCloser + "</tr>";
                                }
                                else if ((list[i].Etd != null || list[i].ExpectedPoissuedDate != null) && (list[i].Eta == null || list[i].Poqty == null))
                                {
                                    var noteetd = list[i].Etd != null ? "-Updated Etd <br>" : "";
                                    var notepodate = list[i].ExpectedPoissuedDate != null ? "-Updated Expected PO Issued Date" : "";

                                    mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" +
                                                        TdOpener + list[i].PreProcId + TdCloser +
                                                        TdOpener + list[i].Idnew + TdCloser +
                                                        TdOpener + list[i].Poid + TdCloser +
                                                        TdOpener + list[i].ItemDescription + TdCloser +
                                                        TdOpener + list[i].ItemNo + TdCloser +
                                                        TdOpener + list[i].Poqty + TdCloser +
                                                        TdOpener + list[i].Eta + TdCloser +
                                                        TdOpener + list[i].ExpectedPoissuedDate + TdCloser +
                                                        TdOpener + list[i].Etd + TdCloser +
                                                        TdOpener + list[i].Grdate + TdCloser +
                                                        TdOpener + list[i].Grqty + TdCloser +
                                                        TdOpener + noteetd + notepodate + TdCloser + "</tr>";
                                }

                            }
                        }
                    }
                    else if (type == "UpdatedRiskAssess")
                    {
                        mailMessage.Body += TextOpenAfterHead + "<br>List Updated Data <br>";//+ TableHeaderRisk;
                        if (riskdata != null)
                        {
                            //if (riskdata.Description != null)
                            //{
                                if (listriskdata.Count > 0)
                                {
                                    mailMessage.Body += "<br> Risks Factors : <br>";
                                    foreach (var item in listriskdata)
                                    {
                                        var tempriskdata = _context.PreProcRisks.OrderBy(x => x.RiskId).FirstOrDefault(x => x.RiskId == item.RiskId);
                                        if (tempriskdata != null)
                                        {
                                            mailMessage.Body += "- " + tempriskdata.RiskName + " (" + item.RiskType + ") : Checked" + "<br>";
                                        }
                                    }
                                }

                                mailMessage.Body += "<br>Detail Assessment : " + "<br>";

                                var noteoverallrisk = overallrisk != "" ? "- Overall Risk : " + overallrisk + "<br>" : "";
                                var notedesc = riskdata.Description != null ? "- Description : " + riskdata.Description + "<br>" : "";
                                var notesource = riskdata.SourceBudget != null ? "- Source Budget : " : "";
                                try
                                {
                                    notesource += _context.MsSourceBudgets.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskdata.SourceBudget)).Type;
                                    notesource += "<br>";
                                }
                                catch (Exception)
                                {
                                    notesource = "";
                                }

                                var noteeststart = riskdata.EstProjectStart != null ? "- Est Project Start : " + riskdata.EstProjectStart.Value.ToShortDateString() + "<br>" : "";
                                var noteestend = riskdata.EstProjectCompletion != null ? "- Est Project End : " + riskdata.EstProjectCompletion.Value.ToShortDateString() + "<br>" : "";
                                var notecompetition = riskdata.Competition != null ? "- Competition : " : "";

                                try
                                {
                                    notecompetition += _context.MsCompetitions.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskdata.Competition)).Type;
                                    notecompetition += "<br>";
                                }
                                catch (Exception)
                                {
                                    notecompetition = "";
                                }

                                var notecontingen = riskdata.ContigencyCost != null ? "- Contingency Cost : " + riskdata.ContigencyCost + "<br>" : "";
                                var noteproctype = riskdata.ProcurementType != null ? "- Procurement Type : " : "";

                                try
                                {
                                    noteproctype += _context.MsProcurementTypes.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskdata.ProcurementType)).Type;
                                    noteproctype += "<br>";
                                }
                                catch (Exception)
                                {
                                    noteproctype = "";
                                }

                                var notecompetitor = riskdata.Competitor != null ? "- Competitor : " + riskdata.Competitor + "<br>" : "";
                                var noteprojsource = riskdata.ProjectSource != null ? "- Project Source : " : "";

                                try
                                {
                                    noteprojsource += _context.MsProjectSources.OrderBy(x => x.Id).FirstOrDefault(x => x.Id == Int32.Parse(riskdata.ProjectSource)).Type;
                                    noteprojsource += "<br>";
                                }
                                catch (Exception)
                                {
                                    noteprojsource = "";
                                }

                                var noteprojcondi = riskdata.ProjectCondition != null ? "- Project Condition : " + riskdata.ProjectCondition + "<br>" : "";
                                var notepredefine = riskdata.PredefinedConditionReqByPrincipal != null ? "- Predefined Condition Requested by Principal/Broker : " + riskdata.PredefinedConditionReqByPrincipal + "<br>" : "";


                                mailMessage.Body += noteoverallrisk + notedesc + notesource + noteeststart + noteestend +
                                                            notecompetition + notecontingen + noteproctype + notecompetitor +
                                                            noteprojsource + noteprojcondi + notepredefine;

                                if (listprojectrev.Count > 0)
                                {
                                    mailMessage.Body += "<br> Revenue Sharing : <br>";
                                    foreach (var item in listprojectrev.Where(x=>x.PrincipalName.Equals("MLPT")))
                                    {
                                        if (item.ProjectRevenue != null)
                                        {
                                            mailMessage.Body += "- " + item.PrincipalName + " : " + item.ProjectRevenue + "%<br>";
                                        }
                                    }

                                    var listprojrevsub = new List<PreProcPrincipalProjectRevenue>();
                                    listprojrevsub.AddRange(listprojectrev.Where(x => !x.PrincipalName.Equals("MLPT")));

                                    mailMessage.Body += "- # Subcon : " + listprojrevsub.Count + "<br>";

                                    mailMessage.Body += TableHeaderProjRev;

                                    foreach (var item in listprojrevsub)
                                    {
                                        if (item.ProjectRevenue != null)
                                        {
                                            var preload = await _context.MsPrincipalPreloads.FirstOrDefaultAsync(x => x.Id == item.Preload);

                                            mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" +
                                                        TdOpener + item.PrincipalName + TdCloser +
                                                        TdOpener + item.ProjectRevenue + TdCloser +
                                                        TdOpener + item.PrincipalScope + TdCloser +
                                                        TdOpener + preload.Type + TdCloser +"</tr>";
                                        }
                                    }
                            }

                            //mailMessage.Body += @"<tr style='mso-yfti-irow:1'>" +
                            //                    TdOpener + riskdata.PreProcId + TdCloser +
                            //                    TdOpener + riskdata.Description + TdCloser +
                            //                    TdOpener + riskdata.SourceBudget + TdCloser +
                            //                    TdOpener + riskdata.EstProjectStart + TdCloser +
                            //                    TdOpener + riskdata.EstProjectCompletion + TdCloser +
                            //                    TdOpener + riskdata.Competition + TdCloser +
                            //                    TdOpener + riskdata.ContigencyCost + TdCloser +
                            //                    TdOpener + riskdata.ProcurementType + TdCloser +
                            //                    TdOpener + riskdata.Competitor + TdCloser +
                            //                    TdOpener + riskdata.ProjectSource + TdCloser +
                            //                    TdOpener + riskdata.ProjectCondition + TdCloser +
                            //                    TdOpener + riskdata.PredefinedConditionReqByPrincipal + TdCloser +
                            //                    TdOpener + notedesc + notesource + TdCloser + "</tr>";

                            
                            //}
                        }
                    }

                    connection.Close();

                    if (type == "UpdatedRemarkFromProc" || type == "UpdatedFromWH" || type == "UpdatedPOFromProc" || type == "UpdatedPODetailFromProc" || type == "UpdatedRiskAssess")
                    {
                        mailMessage.Body += TextClosedWithTable;
                    }

                    var urlapp = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Url App");

                    mailMessage.Body += "<br>Please click the hyperlink below for further details: <br><br> <a  class='btn-down-spec' style='font-size: 12px; padding: 7px!important;margin-right: 5px;margin-left: 5px;' href='" + urlapp.AlphaNumericValue + "/Preprocs/ViewDetail/" + proc.Id + "'>Click here to open this item</a> ";

                    mailMessage.Body += TextClosed;

                    try
                    {
                        using (var smtpClient = new SmtpClient())
                        {
                            smtpClient.Host = "10.10.62.18"; //production 64.229
                            smtpClient.Send(mailMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (SqlException e)
            {
                //break;
            }

            //return Ok();
            return data;
        }

        [HttpPost]
        public async Task<IActionResult> PostSaveItemDetail(string[] principal, string[] detailid, string[] nkpcode, string[] itemdesc, string[] insurance, string[] warranty, string[] specrks, string[] vendorquot, string[] vendor, string[] surduk, string[] surdukatt, string[] qty, string[] curr, string[] unitcogs, string[] unitrev, string[] totcogs, string[] totrev, int procid, string[] no, string[] risk, string[] presalesreview, string[] amreview, string[] presalescomment, string[] amcomment, string[] existingspecrks, string[] existingsurduk, string[] existingvendorquot, string[] remark)
        {
            var usernow = getUserNow();

            PreProcDetail preProcDetail;

            var linecount = 1;

            var preproc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);
            var doctypespec = await _context.MsDocTypes.FirstOrDefaultAsync(x => x.DocType == "SPC");
            var doctypevendorquot = await _context.MsDocTypes.FirstOrDefaultAsync(x => x.DocType == "VQU");
            var doctypesurduk = await _context.MsDocTypes.FirstOrDefaultAsync(x => x.DocType == "SRD");

            for (var i = 0; i < principal.Length; i++)
            {
                if (detailid[i] != null && detailid[i] != "")
                {
                    var detail = Int32.Parse(detailid[i]);

                    if (detail != 0)
                    {
                        var tempdetail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detail);
                        var itemdescnkp = await _context.VwGetNkpnames.FirstOrDefaultAsync(x => x.NkpcodeLinkTitle == nkpcode[i]);

                        tempdetail.Id = procid;
                        tempdetail.Principle = principal[i];
                        tempdetail.Nkpcode = nkpcode[i];
                        tempdetail.ItemDesc = itemdescnkp != null ? itemdescnkp.Nkpname : null;
                        tempdetail.Insurance = insurance[i].Trim() == "Y" ? true : false;
                        tempdetail.Warranty = Int32.Parse(warranty[i]);
                        tempdetail.Qty = Int32.Parse(qty[i]);
                        surduk[i] = surduk[i] == null ? null : surduk[i].Trim();
                        tempdetail.SurDukCheck = surduk[i] == "N" ? false : true;
                        tempdetail.Spec = specrks[i] != null && specrks[i] != "" ? "Attach" : null;
                        tempdetail.SurDuk = surdukatt[i] != null && surdukatt[i] != "" ? "Attach" : null;
                        tempdetail.VendorQuotation = vendorquot[i] != null && vendorquot[i] != "" ? "Attach" : null;
                        tempdetail.Idnew = no[i];
                        tempdetail.Currency = curr[i];
                        tempdetail.UnitCogs = unitcogs[i] != null ? double.Parse(unitcogs[i]) : null;
                        tempdetail.UnitPrice = unitrev[i] != null ? double.Parse(unitrev[i]) : null;
                        tempdetail.TotalCogs = totcogs[i];
                        tempdetail.TotalPrice = totrev[i];
                        tempdetail.Risk = risk[i];
                        tempdetail.PresalesReview = presalesreview[i];
                        tempdetail.AmmanagerApproval = amreview[i];
                        tempdetail.RemarkOnItem = remark[i];
                        tempdetail.Vendor = vendor[i];
                        
                        if (no[i].Contains("."))
                        {
                            var splitno = no[i].Split('.');
                            tempdetail.FirstId = Int32.Parse(splitno[0]);
                            tempdetail.SecondId = splitno[1] != "" ? Int32.Parse(splitno[1]) : null;
                        }
                        else
                        {
                            tempdetail.FirstId = Int32.Parse(no[i]);
                        }

                        await _context.SaveChangesAsync();

                        var tempapprovalpre = await _context.PreProcApprovalHistories.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == procid && x.ApproverRole == "Presales" && x.DetailId == detail);
                        var tempapprovalam = await _context.PreProcApprovalHistories.OrderByDescending(x => x.Id).FirstOrDefaultAsync(x => x.PreProcId == procid && x.ApproverRole == "AM Manager" && x.DetailId == detail);

                        var commenttype = _context.MsCommentTypes.FirstOrDefault(x => x.Type == "Detail");

                        if (presalesreview[i] != null)
                        {
                            if ( (tempapprovalpre == null && presalesreview[i] != null) || (tempapprovalpre != null && !tempapprovalpre.Status.Contains(presalesreview[i])))
                        {
                            PreProcApprovalHistory approvalhistory = new PreProcApprovalHistory();

                            approvalhistory.ApprovalDate = DateTime.Now;
                            approvalhistory.Remark = presalescomment[i];
                            approvalhistory.Status = presalesreview[i] + ": Line " + linecount;
                            approvalhistory.DetailId = detail;
                            approvalhistory.DomainName = User.Identity.Name;
                            approvalhistory.ApproverName = usernow;
                            approvalhistory.ApproverRole = "Presales";
                            approvalhistory.PreProcId = procid;
                            approvalhistory.Type = commenttype.Id;

                            _context.Add(approvalhistory);

                            await _context.SaveChangesAsync();
                        }
                        }

                        if (amreview[i] != null)
                        {
                            
                            if ((tempapprovalam == null && amreview[i] != null) || (tempapprovalam != null && !tempapprovalam.Status.Contains(amreview[i])))
                        {
                            PreProcApprovalHistory approvalhistory = new PreProcApprovalHistory();

                            approvalhistory.ApprovalDate = DateTime.Now;
                            approvalhistory.Remark = amcomment[i];
                            approvalhistory.Status = amreview[i] + ": Line " + linecount;
                            approvalhistory.DetailId = detail;
                            approvalhistory.DomainName = User.Identity.Name;
                            approvalhistory.ApproverName = usernow;
                            approvalhistory.ApproverRole = "AM Manager";
                            approvalhistory.PreProcId = procid;
                            approvalhistory.Type = commenttype.Id;

                            _context.Add(approvalhistory);

                            await _context.SaveChangesAsync();
                        }
                        }

                        if (specrks[i] != null && specrks[i] != "")
                        {
                            var split = specrks[i].Split(";");

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempspec = await _context.PreProcHeaderAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    PreProcFileAttachment fileattach = new PreProcFileAttachment();

                                    fileattach.FileName = tempspec.FileName;
                                    fileattach.StrictedFileName = tempspec.StrictedFileName;
                                    fileattach.Folder = preproc.PresalesId;
                                    fileattach.PreProcId = procid;
                                    fileattach.DocumentTypeId = doctypespec.Id;
                                    fileattach.Bastid = detail;
                                    fileattach.SpecAttachId = detail;
                                    fileattach.HasBeenUploaded = true;

                                    _context.Add(fileattach);
                                    await _context.SaveChangesAsync();

                                    _context.Remove(tempspec);
                                    await _context.SaveChangesAsync();
                                }
                            }
                        }

                        if (vendorquot[i] != null && vendorquot[i] != "")
                        {
                            var split = vendorquot[i].Split(";");

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempvendorquot = await _context.PreProcHeaderAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    PreProcFileAttachment fileattach = new PreProcFileAttachment();

                                    fileattach.FileName = tempvendorquot.FileName;
                                    fileattach.StrictedFileName = tempvendorquot.StrictedFileName;
                                    fileattach.Folder = preproc.PresalesId;
                                    fileattach.PreProcId = procid;
                                    fileattach.DocumentTypeId = doctypevendorquot.Id;
                                    fileattach.Bastid = detail;
                                    fileattach.VendorQuoAttachId = detail;
                                    fileattach.HasBeenUploaded = true;

                                    _context.Add(fileattach);
                                    await _context.SaveChangesAsync();

                                    _context.Remove(tempvendorquot);
                                    await _context.SaveChangesAsync();
                                }
                            }
                        }

                        if (surdukatt[i] != null && surdukatt[i] != "")
                        {
                            var split = surdukatt[i].Split(";");

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempsurduk = await _context.PreProcHeaderAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    PreProcFileAttachment fileattach = new PreProcFileAttachment();

                                    fileattach.FileName = tempsurduk.FileName;
                                    fileattach.StrictedFileName = tempsurduk.StrictedFileName;
                                    fileattach.Folder = preproc.PresalesId;
                                    fileattach.PreProcId = procid;
                                    fileattach.DocumentTypeId = doctypesurduk.Id;
                                    fileattach.Bastid = detail;
                                    fileattach.SurdukAttachId = detail;
                                    fileattach.HasBeenUploaded = true;

                                    _context.Add(fileattach);
                                    await _context.SaveChangesAsync();

                                    _context.Remove(tempsurduk);
                                    await _context.SaveChangesAsync();
                                }
                            }
                        }

                        if (existingspecrks[i] != null)
                        {
                            var split = existingspecrks[i].Split(";");
                            var listid = await _context.PreProcFileAttachments.Where(x => x.Bastid == detail && x.DocumentTypeId == 11).ToListAsync();
                            var delete = "";

                            if (listid != null)
                            {
                                foreach (var item in listid)
                                {
                                    delete += item.Id + ";";
                                }
                            }

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempexistspec = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x=>x.Id == Int32.Parse(item));

                                    if (tempexistspec != null)
                                    {
                                        delete = delete.Replace(item + ";", "");
                                    }

                                }

                            }

                            if (delete != null)
                            {
                                split = delete.Split(";");

                                foreach (var item in split)
                                {
                                    if (item != "")
                                    {
                                        var tempexistspec = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                        if (tempexistspec != null)
                                        {
                                            tempexistspec.HasBeenUploaded = false;

                                            await _context.SaveChangesAsync();
                                        }

                                    }

                                }
                            }
                            

                        }

                        if (existingvendorquot[i] != null)
                        {
                            var split = existingvendorquot[i].Split(";");
                            var listid = await _context.PreProcFileAttachments.Where(x => x.Bastid == detail && x.DocumentTypeId == 13).ToListAsync();
                            var delete = "";

                            if (listid != null)
                            {
                                foreach (var item in listid)
                                {
                                    delete += item.Id + ";";
                                }
                            }

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempexistvendorquot = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    if (tempexistvendorquot != null)
                                    {
                                        delete = delete.Replace(item + ";", "");
                                    }

                                }

                            }

                            if (delete != null)
                            {
                                split = delete.Split(";");

                                foreach (var item in split)
                                {
                                    if (item != "")
                                    {
                                        var tempexistvendorquot = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                        if (tempexistvendorquot != null)
                                        {
                                            tempexistvendorquot.HasBeenUploaded = false;

                                            await _context.SaveChangesAsync();
                                        }

                                    }

                                }
                            }


                        }

                        if (existingsurduk[i] != null)
                        {
                            var split = existingsurduk[i].Split(";");
                            var listid = await _context.PreProcFileAttachments.Where(x => x.Bastid == detail && x.DocumentTypeId == 12).ToListAsync();
                            var delete = "";

                            if (listid != null)
                            {
                                foreach (var item in listid)
                                {
                                    delete += item.Id + ";";
                                }
                            }

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempexistsurduk = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    if (tempexistsurduk != null)
                                    {
                                        delete = delete.Replace(item + ";", "");
                                    }

                                }

                            }

                            if (delete != null)
                            {
                                split = delete.Split(";");

                                foreach (var item in split)
                                {
                                    if (item != "")
                                    {
                                        var tempexistsurduk = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                        if (tempexistsurduk != null)
                                        {
                                            tempexistsurduk.HasBeenUploaded = false;

                                            await _context.SaveChangesAsync();
                                        }

                                    }

                                }
                            }


                        }
                    }
                    else
                    {
                        var itemdescnkp = await _context.VwGetNkpnames.FirstOrDefaultAsync(x => x.NkpcodeLinkTitle == nkpcode[i]);
                        
                        preProcDetail = new PreProcDetail();
                        preProcDetail.Id = procid;
                        preProcDetail.Principle = principal[i];
                        preProcDetail.Nkpcode = nkpcode[i];
                        preProcDetail.ItemDesc = itemdescnkp != null ? itemdescnkp.Nkpname : null;
                        preProcDetail.Insurance = insurance[i].Trim() == "Y" ? true : false;
                        preProcDetail.Warranty = Int32.Parse(warranty[i]);
                        preProcDetail.Qty = Int32.Parse(qty[i]);
                        surduk[i] = surduk[i] == null ? null : surduk[i].Trim();
                        preProcDetail.SurDukCheck = surduk[i] == "N" ? false : true;
                        preProcDetail.Spec = specrks[i] != null && specrks[i] != "" ? "Attach" : null;
                        preProcDetail.SurDuk = surdukatt[i] != null && surdukatt[i] != "" ? "Attach" : null;
                        preProcDetail.Idnew = no[i];
                        preProcDetail.Currency = curr[i];
                        preProcDetail.UnitCogs = unitcogs[i] != null ? double.Parse(unitcogs[i]) : null;
                        preProcDetail.UnitPrice = unitrev[i] != null ? double.Parse(unitrev[i]) : null;
                        preProcDetail.TotalCogs = totcogs[i];
                        preProcDetail.TotalPrice = totrev[i];
                        preProcDetail.RemarkOnItem = remark[i];
                        
                        if (no[i].Contains("."))
                        {
                            var splitno = no[i].Split('.');
                            preProcDetail.FirstId = Int32.Parse(splitno[0]);
                            preProcDetail.SecondId = splitno[1] != "" ? Int32.Parse(splitno[1]) : null;
                        }
                        else
                        {
                            preProcDetail.FirstId = Int32.Parse(no[i]);
                        }
                        _context.Add(preProcDetail);
                        await _context.SaveChangesAsync();

                        if (specrks[i] != null && specrks[i] != "")
                        {
                            var split = specrks[i].Split(";");

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempspec = await _context.PreProcHeaderAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    PreProcFileAttachment fileattach = new PreProcFileAttachment();

                                    fileattach.FileName = tempspec.FileName;
                                    fileattach.StrictedFileName = tempspec.StrictedFileName;
                                    fileattach.Folder = preproc.PresalesId;
                                    fileattach.PreProcId = procid;
                                    fileattach.DocumentTypeId = doctypespec.Id;
                                    fileattach.Bastid = preProcDetail.DetailId;
                                    fileattach.SpecAttachId = preProcDetail.DetailId;
                                    fileattach.HasBeenUploaded = true;

                                    _context.Add(fileattach);
                                    await _context.SaveChangesAsync();

                                    _context.Remove(tempspec);
                                    await _context.SaveChangesAsync();
                                }
                            }
                        }
                        if (surdukatt[i] != null && surdukatt[i] != "")
                        {
                            var split = surdukatt[i].Split(";");

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempsurduk = await _context.PreProcHeaderAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    PreProcFileAttachment fileattach = new PreProcFileAttachment();

                                    fileattach.FileName = tempsurduk.FileName;
                                    fileattach.StrictedFileName = tempsurduk.StrictedFileName;
                                    fileattach.Folder = preproc.PresalesId;
                                    fileattach.PreProcId = procid;
                                    fileattach.DocumentTypeId = doctypesurduk.Id;
                                    fileattach.Bastid = preProcDetail.DetailId;
                                    fileattach.SurdukAttachId = preProcDetail.DetailId;
                                    fileattach.HasBeenUploaded = true;

                                    _context.Add(fileattach);
                                    await _context.SaveChangesAsync();

                                    _context.Remove(tempsurduk);
                                    await _context.SaveChangesAsync();
                                }
                            }
                        }
                        if (existingspecrks[i] != null)
                        {
                            var split = existingspecrks[i].Split(";");
                            var listid = await _context.PreProcFileAttachments.Where(x => x.Bastid == preProcDetail.DetailId && x.DocumentTypeId == 11).ToListAsync();
                            var delete = "";

                            if (listid != null)
                            {
                                foreach (var item in listid)
                                {
                                    delete += item.Id + ";";
                                }
                            }

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempexistspec = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    if (tempexistspec != null)
                                    {
                                        delete = delete.Replace(item + ";", "");
                                    }

                                }

                            }

                            if (delete != null)
                            {
                                split = delete.Split(";");

                                foreach (var item in split)
                                {
                                    if (item != "")
                                    {
                                        var tempexistspec = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                        if (tempexistspec != null)
                                        {
                                            tempexistspec.HasBeenUploaded = false;

                                            await _context.SaveChangesAsync();
                                        }

                                    }

                                }
                            }


                        }
                        if (existingsurduk[i] != null)
                        {
                            var split = existingsurduk[i].Split(";");
                            var listid = await _context.PreProcFileAttachments.Where(x => x.Bastid == preProcDetail.DetailId && x.DocumentTypeId == 12).ToListAsync();
                            var delete = "";

                            if (listid != null)
                            {
                                foreach (var item in listid)
                                {
                                    delete += item.Id + ";";
                                }
                            }

                            foreach (var item in split)
                            {
                                if (item != "")
                                {
                                    var tempexistsurduk = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                    if (tempexistsurduk != null)
                                    {
                                        delete = delete.Replace(item + ";", "");
                                    }

                                }

                            }

                            if (delete != null)
                            {
                                split = delete.Split(";");

                                foreach (var item in split)
                                {
                                    if (item != "")
                                    {
                                        var tempexistsurduk = await _context.PreProcFileAttachments.FirstOrDefaultAsync(x => x.Id == Int32.Parse(item));

                                        if (tempexistsurduk != null)
                                        {
                                            tempexistsurduk.HasBeenUploaded = false;

                                            await _context.SaveChangesAsync();
                                        }

                                    }

                                }
                            }


                        }

                    }

                    linecount++;
                }
            }
            //return RedirectToAction(nameof(Index));
            return Ok();
        }

        [HttpPost]
        public async Task<IActionResult> PostEditItem(PreProcGeneralInfo2 model, IFormFile postedfilerfp, IFormFile postedfileboq, IFormFile postedfilequot, IFormFile postedfilegp, IFormFile postedfileproposal, IFormFile postedfileothers, bool NeedEsc, int txtBillReason, string txtOtherReason, IFormFile postedfilepmorfp, IFormFile postedfilepmoboq, IFormFile postedfilepmoquot, IFormFile postedfilepmoproposal, IFormFile postedfilepmoother, IFormFile postedfilepmonego, IFormFile postedfilepmospk, decimal txtValuePM, decimal txtValuePC, int txtMandaysPM, int txtMandaysPC)
        {
            getUserNow();
            #region COMMENT
            //PreProcGeneralInfo2 preProcGI = new PreProcGeneralInfo2();//= new preProcGI();
            //PreProcHeaderAttachment preProcAttach = new PreProcHeaderAttachment();//= new preProcAttach();

            //if (model.Stage == "RiskAssessment")
            //{

            //}
            //else if (model.Stage == "Presales")
            //{
            //    preProcGI.Stage = model.Stage;
            //    preProcGI.RfpId = model.RfpId;
            //    preProcGI.ProposalId = model.ProposalId;
            //    preProcGI.QuotationId = model.QuotationId;
            //    var tempCust = model.CustomerCode.Split(";");
            //    preProcGI.Customer = tempCust[0];
            //    preProcGI.CustomerCode = tempCust[1];
            //    preProcGI.Am = model.Am;
            //    preProcGI.PropMgr = model.PropMgr;
            //    preProcGI.TenderAdm = model.TenderAdm;
            //    preProcGI.ProposalSubmissionDate = model.ProposalSubmissionDate;
            //    preProcGI.QuotationDeadline = model.QuotationDeadline;
            //    preProcGI.Hedging = model.Hedging;
            //    preProcGI.CreatedBy = "UserTesting1";
            //    preProcGI.CreationDate = DateTime.Now;
            //    preProcGI.Created = DateTime.Now;
            //    preProcGI.NeedWarranty = model.NeedWarranty;
            //    preProcGI.PresalesId = model.PresalesId;
            //    preProcGI.Pid = model.PresalesId;
            //    preProcGI.ProjectValueIdr = model.ProjectValueIdr;
            //    preProcGI.ProjectValueUsd = model.ProjectValueUsd;
            //    preProcGI.EstGpgeneral = model.EstGpgeneral;
            //    preProcGI.ProjectType = model.ProjectType;
            //    preProcGI.Description = model.Description;
            //    preProcGI.PicProc = model.PicProc;

            //}
            //else
            //{

            //}
            //_context.Add(preProcGI);
            //await _context.SaveChangesAsync();
            //if (postedfilerfp != null)
            //{
            //    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + model.PresalesId);

            //    //create folder if not exist
            //    if (!Directory.Exists(path))
            //        Directory.CreateDirectory(path);


            //    string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
            //                                                 //+
            //                                                 postedfilerfp.FileName);

            //    var doctype = from t in _context.MsDocTypes
            //                  select t;

            //    var tempdoctype = doctype.Where(t => t.DocType.Equals("RFP")).FirstOrDefaultAsync();
            //    using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
            //    {
            //        postedfilerfp.CopyTo(stream);
            //    }
            //    preProcAttach.Folder = model.PresalesId;
            //    preProcAttach.FileName = postedfilerfp.FileName;
            //    preProcAttach.StrictedFileName = fileNameWithPath;
            //    preProcAttach.ModifiedBy = "tes";
            //    preProcAttach.UploadedFrom = "PREPROC";
            //    preProcAttach.HasBeenUploaded = true;
            //    preProcAttach.DocumentTypeId = tempdoctype.Result.Id;
            //    preProcAttach.PreProcId = preProcGI.Id;
            //}

            //_context.Add(preProcAttach);
            //await _context.SaveChangesAsync();
            #endregion

            var detail = await _context.PreProcDetails.Where(x => x.Id == model.Id).ToListAsync();
            var pmocheck = await _context.PreProcPmochecklists.FirstOrDefaultAsync(x => x.PreProcId == model.Id);

            PreProcDetail preProcDetail;

            if (detail.Count > 0)
            {
                foreach (var item in detail)
                {
                    if (item.DetailId == null)
                    {
                        preProcDetail = new PreProcDetail();
                        //preProcDetail.Id = item.Id;
                        //preProcDetail.Principle = item.Principle;
                        //preProcDetail.Nkpcode = item.Nkpcode;
                        //preProcDetail.ItemDesc = item.ItemDesc;
                        //preProcDetail.Insurance = item.Insurance;
                        //preProcDetail.Warranty = item.Warranty;
                        //preProcDetail.Qty = item.Qty;
                        //preProcDetail.SurDukCheck = item.SurDukCheck;
                        //preProcDetail.Idnew = item.Idnew;
                        //preProcDetail.SpecAttachId = item.SpecAttachId;
                        //preProcDetail.SurdukAttachId = item.SurdukAttachId;
                        //preProcDetail.Currency = item.Currency;
                        //preProcDetail.UnitCogs = item.UnitCogs;
                        //preProcDetail.UnitPrice = item.UnitPrice;
                        //preProcDetail.TotalCogs = item.TotalCogs;
                        //preProcDetail.TotalPrice = item.TotalPrice;
                        //preProcDetail.Vendor = item.Vendor;
                        //preProcDetail.VendorQuotation = item.VendorQuotation;
                        //preProcDetail.ExpectedPoissuedDate = item.ExpectedPoissuedDate;
                        //preProcDetail.Mlptpono = item.Mlptpono;
                        //preProcDetail.Mlptpodate = item.Mlptpodate;
                        //preProcDetail.Mlptpo = item.Mlptpo;
                        //preProcDetail.Eta = item.Eta;
                        //preProcDetail.Grdate = item.Grdate;
                        //preProcDetail.Grqty = item.Grqty;
                        //preProcDetail.Grbacklog = item.Grbacklog;
                        //preProcDetail.PresalesReview = item.PresalesReview;
                        //preProcDetail.AmmanagerApproval = item.AmmanagerApproval;
                        preProcDetail.RemarkOnItem = item.RemarkOnItem;
                        //preProcDetail.ProcessedPresalesReview = item.ProcessedPresalesReview;
                        //preProcDetail.ProcessedAmreview = item.ProcessedAmreview;
                        //preProcDetail.PonumberId = item.PonumberId;
                        //preProcDetail.Nkpcode = item.Nkpcode;
                        //preProcDetail.ProductManager = item.ProductManager;
                        //preProcDetail.Warranty = item.Warranty;
                        //preProcDetail.Insurance = item.Insurance;
                        //preProcDetail.Risk = item.Risk;
                        //preProcDetail.AuditDetailId = item.AuditDetailId;
                        //preProcDetail.Idnew = item.Idnew;
                        //preProcDetail.Grn = item.Grn;
                        //preProcDetail.Pidfull = item.Pidfull;
                        //preProcDetail.FirstId = item.FirstId;
                        //preProcDetail.SecondId = item.SecondId;
                        _context.Add(preProcDetail);
                        await _context.SaveChangesAsync();
                    }

                    //var detupdate = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.Id == model.Id && x.DetailId == item.DetailId);

                    //if (detupdate != null)
                    //{
                    //    detupdate.Principle = item.Principle;
                    //    detupdate.Nkpcode = item.Nkpcode;
                    //    detupdate.ItemDesc = item.ItemDesc;
                    //    detupdate.Insurance = item.Insurance;
                    //    detupdate.Warranty = item.Warranty;
                    //    detupdate.Qty = item.Qty;
                    //    detupdate.SurDukCheck = item.SurDukCheck;
                    //    detupdate.Idnew = item.Idnew;
                    //    detupdate.SpecAttachId = item.SpecAttachId;
                    //    detupdate.SurdukAttachId = item.SurdukAttachId;
                    //    detupdate.Currency = item.Currency;
                    //    detupdate.UnitCogs = item.UnitCogs;
                    //    detupdate.UnitPrice = item.UnitPrice;
                    //    detupdate.TotalCogs = item.TotalCogs;
                    //    detupdate.TotalPrice = item.TotalPrice;
                    //    detupdate.Vendor = item.Vendor;
                    //    detupdate.VendorQuotation = item.VendorQuotation;
                    //    detupdate.ExpectedPoissuedDate = item.ExpectedPoissuedDate;
                    //    detupdate.Mlptpono = item.Mlptpono;
                    //    detupdate.Mlptpodate = item.Mlptpodate;
                    //    detupdate.Mlptpo = item.Mlptpo;
                    //    detupdate.Eta = item.Eta;
                    //    detupdate.Grdate = item.Grdate;
                    //    detupdate.Grqty = item.Grqty;
                    //    detupdate.Grbacklog = item.Grbacklog;
                    //    detupdate.PresalesReview = item.PresalesReview;
                    //    detupdate.AmmanagerApproval = item.AmmanagerApproval;
                    //    detupdate.RemarkOnItem = item.RemarkOnItem;
                    //    detupdate.ProcessedPresalesReview = item.ProcessedPresalesReview;
                    //    detupdate.ProcessedAmreview = item.ProcessedAmreview;
                    //    detupdate.PonumberId = item.PonumberId;
                    //    detupdate.Nkpcode = item.Nkpcode;
                    //    detupdate.ProductManager = item.ProductManager;
                    //    detupdate.Warranty = item.Warranty;
                    //    detupdate.Insurance = item.Insurance;
                    //    detupdate.Risk = item.Risk;
                    //    detupdate.AuditDetailId = item.AuditDetailId;
                    //    detupdate.Idnew = item.Idnew;
                    //    detupdate.Grn = item.Grn;
                    //    detupdate.Pidfull = item.Pidfull;
                    //    detupdate.FirstId = item.FirstId;
                    //    detupdate.SecondId = item.SecondId;
                    //    await _context.SaveChangesAsync();
                    //}

                    //var fileattach = await _context.PreProcFileAttachments.Where(x => x.Bastid == item.DetailId).ToListAsync();

                    //foreach (var item1 in fileattach)
                    //{
                    //    item1.Bastid = item.DetailId;
                    //    await _context.SaveChangesAsync();
                    //}

                    //if (model.Stage == "Presales")
                    //{
                    //    var approvalhistpre = await _context.PreProcApprovalHistories.FirstOrDefaultAsync(x => x.DetailId == item.DetailId && x.Status.Contains("temp") && x.ApproverRole == "Presales");

                    //    var approvalhist = await _context.PreProcApprovalHistories.FirstOrDefaultAsync(x => x.DetailId == item.DetailId && x.Status.Contains("temp") && x.ApproverRole != "Presales");

                    //    if (approvalhistpre != null)
                    //    {
                    //        var tempstat = approvalhistpre.Status.Split(";");

                    //        approvalhistpre.Status = item.PresalesReview + " : " + tempstat != null && tempstat.Length > 1 ? tempstat[1] : null;
                    //        approvalhistpre.ApprovalDate = DateTime.Now;
                    //        approvalhistpre.DetailId = item.DetailId;
                    //        await _context.SaveChangesAsync();

                    //    }

                    //    if (approvalhist != null)
                    //    {
                    //        var tempstat = approvalhist.Status.Split(";");

                    //        approvalhist.Status = item.AmmanagerApproval + " : " + tempstat != null && tempstat.Length > 1 ? tempstat[1] : null;
                    //        approvalhist.ApprovalDate = DateTime.Now;
                    //        approvalhist.DetailId = item.DetailId;
                    //        await _context.SaveChangesAsync();

                    //    }
                    //}
                    

                }

                
                //foreach (var item in detail)
                //{
                //    _context.Remove(item);
                //    await _context.SaveChangesAsync();
                //}
            }

            var preProcGI = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == model.Id);

            preProcGI.Stage = model.Stage;
            preProcGI.RfpId = model.RfpId;
            preProcGI.ProposalId = model.ProposalId;
            preProcGI.QuotationId = model.QuotationId;
            preProcGI.Customer = model.Customer;
            preProcGI.Am = model.Am;
            preProcGI.PropMgr = model.PropMgr;
            preProcGI.TenderAdm = model.TenderAdm;
            preProcGI.ProposalSubmissionDate = model.ProposalSubmissionDate;
            preProcGI.QuotationDeadline = model.QuotationDeadline;
            preProcGI.Hedging = model.Hedging;
            preProcGI.NeedWarranty = model.NeedWarranty;
            preProcGI.PresalesId = model.PresalesId;
            preProcGI.Pid = model.PresalesId;
            if (model.Stage == "Delivery")
            {
                preProcGI.PresalesId = model.Pid;
                preProcGI.Pid = model.Pid;
            }
            preProcGI.ProjectValueIdr = model.ProjectValueIdr;
            preProcGI.ProjectValueUsd = model.ProjectValueUsd;
            preProcGI.EstGpgeneral = model.EstGpgeneral;
            preProcGI.ProjectType = model.ProjectType;
            preProcGI.Description = model.Description;
            preProcGI.PicProc = model.PicProc;
            preProcGI.Pc = model.Pc;
            preProcGI.Pm = model.Pm;
            preProcGI.ProjectSupport = model.ProjectSupport;
            preProcGI.EmailCommitteeDefault = model.EmailCommitteeDefault;

            await _context.SaveChangesAsync();

            var existbillreason = await _context.PreProcPendingBillingReasons.FirstOrDefaultAsync(x => x.PreProcId == model.Id);

            if (existbillreason != null && txtBillReason > 0)
            {
                int flag = 0;

                if (existbillreason.NeedEscalation != NeedEsc)
                {
                    flag = 1;
                }
                else if (existbillreason.EscalationReasonId != txtBillReason)
                {
                    flag = 1;
                }
                else if (existbillreason.ReasonPic != model.PicProc)
                {
                    flag = 1;
                }
                else if (existbillreason.OtherReason != txtOtherReason)
                {
                    flag = 1;
                }

                if (flag == 1)
                {
                    existbillreason.NeedEscalation = NeedEsc;
                    existbillreason.EscalationReasonId = txtBillReason;
                    existbillreason.ReasonUpdateDate = DateTime.Now;
                    existbillreason.ReasonPic = model.PicProc;
                    existbillreason.OtherReason = txtOtherReason;

                    await _context.SaveChangesAsync();
                }

            }
            else if (txtBillReason > 0)
            {
                PreProcPendingBillingReason billreason = new PreProcPendingBillingReason();

                billreason.PreProcId = model.Id;
                billreason.NeedEscalation = NeedEsc;
                billreason.EscalationReasonId = txtBillReason;
                billreason.ReasonCreatedDate = DateTime.Now;
                billreason.ReasonPic = model.PicProc;
                billreason.OtherReason = txtOtherReason;

                _context.Add(billreason);
                await _context.SaveChangesAsync();
            }

            PreProcHeaderAttachment preProcAttach;

            if (postedfilerfp != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilerfp.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("RFP")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilerfp.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilerfp.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfileboq != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfileboq.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("BOQ")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfileboq.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfileboq.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilequot != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilequot.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("QUO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilequot.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilequot.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilegp != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilegp.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("GPJ")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilegp.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilegp.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfileproposal != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfileproposal.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("PRO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfileproposal.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfileproposal.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfileothers != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfileothers.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("OTH")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfileothers.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfileothers.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }

            if (postedfilepmospk != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmospk.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("SPK")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmospk.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmospk.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilepmorfp != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmorfp.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("RFP")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmorfp.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmorfp.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilepmoboq != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmoboq.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("BOQ")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmoboq.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmoboq.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilepmoquot != null)
            {
                preProcAttach = new PreProcHeaderAttachment();
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmoquot.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("QUO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmoquot.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmoquot.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilepmonego != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmonego.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("NMM")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmonego.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmonego.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilepmoproposal != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmoproposal.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("PRO")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmoproposal.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmoproposal.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }
            if (postedfilepmoother != null)
            {
                preProcAttach = new PreProcHeaderAttachment();

                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preProcGI.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string fileNameWithPath = Path.Combine(path, //DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_"
                                                             //+
                                                                postedfilepmoother.FileName);

                var doctype = from t in _context.MsDocTypes
                              select t;

                var tempdoctype = await doctype.Where(t => t.DocType.Equals("OTH")).FirstOrDefaultAsync();
                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    postedfilepmoother.CopyTo(stream);
                }
                preProcAttach.Folder = preProcGI.PresalesId;
                preProcAttach.FileName = postedfilepmoother.FileName;
                preProcAttach.StrictedFileName = fileNameWithPath;
                preProcAttach.ModifiedBy = User.Identity.Name;
                preProcAttach.UploadedFrom = "PREPROC";
                preProcAttach.HasBeenUploaded = true;
                preProcAttach.DocumentTypeId = tempdoctype.Id;
                preProcAttach.PreProcId = preProcGI.Id;

                _context.Add(preProcAttach);
                await _context.SaveChangesAsync();
            }

            if (pmocheck != null)
            {
                pmocheck.Pmvalue = txtValuePM != 0 ? txtValuePM : null;
                pmocheck.Pcvalue = txtValuePC != 0 ? txtValuePC : null;
                pmocheck.Pmmandays = txtMandaysPM != 0 ? txtMandaysPM : null;
                pmocheck.Pcmandays = txtMandaysPC != 0 ? txtMandaysPC : null;

                await _context.SaveChangesAsync();
            }

            //return RedirectToAction(nameof(Index));
            return RedirectToAction("EditDetail", "PreProcs",
                new{
                    id = model.Id
                });
        }

        [HttpPost]
        public async Task<PreProcDetail> PostEditRemarkItem(int preprocid, int detailid, string remark)
        {
            getUserNow();
            var data = new PreProcDetail();
            var listdata = new List<PreProcDetail>();

            if (detailid != null)
            {
                PreProcDetail temp = new PreProcDetail();
                var tempData = await _context.PreProcDetails.OrderByDescending(x=>x.DetailId).FirstOrDefaultAsync(x => x.DetailId == detailid && x.Id == preprocid);

                if (tempData != null)
                {
                    var split = tempData.Idnew.Split(".");

                    temp = tempData;
                    temp.RemarkOnItem = remark;
                    temp.Idnew = "EDIT";
                    temp.DetailId = 0;


                    if (split.Length > 1)
                    {
                        temp.SecondId = Int32.Parse(split[1]);
                    }
                    temp.FirstId = Int32.Parse(split[0]);

                    _context.Add(temp);
                    await _context.SaveChangesAsync();
                    data = tempData;
                    listdata.Add(tempData);
                }
            }

            await SendNotifByType("", preprocid, "UpdatedRemarkFromProc", new List<PreProcPodatum>(), listdata, new PreProcRiskAssesmentDatum(), "", new List<PreProcRiskDatum>(), new List<PreProcPrincipalProjectRevenue>());

            return data;
        }

        [HttpPost]
        public async Task<string> SaveAttachSpec(List<IFormFile> fileUploadspec, string preprocid, int detailid)
        {
            getUserNow();
            var data = new PreProcDetail();
            PreProcHeaderAttachment fileAttachment;
            PreProcDetail tempDetail;
            string idattach = "";

            //if (detailid == 0)
            //{
            //    tempDetail = new PreProcDetail();

            //    tempDetail.Id = Int32.Parse(preprocid);
            //    tempDetail.Spec = "Attach";
            //    //tempDetail.SpecAttachId = fileAttachment.Id;
            //    _context.Add(tempDetail);
            //    await _context.SaveChangesAsync();

            //    data = tempDetail;
            //}
            //else
            //{
            //    tempDetail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid);

            //    if (tempDetail != null)
            //    {
            //        tempDetail.Spec = "Attach";
            //        //tempDetail.SpecAttachId = fileAttachment.Id;
            //        await _context.SaveChangesAsync();

            //        data = tempDetail;
            //    }
            //}

            foreach (var item in fileUploadspec)
            {
                if (fileUploadspec != null)
                {
                    var preprocdet = await _context.PreProcDetails.Where(x => x.Id == detailid).FirstOrDefaultAsync();
                    var preproc = await _context.PreProcGeneralInfo2s.Where(x => x.Id == Int32.Parse(preprocid)).FirstOrDefaultAsync();

                    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                    //create folder if not exist
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);


                    string filename = DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_" + item.FileName;

                    string fileNameWithPath = Path.Combine(path, filename);

                    var doctype = from t in _context.MsDocTypes
                                  select t;

                    var tempdoctype = await doctype.Where(t => t.DocType.Equals("SPC")).FirstOrDefaultAsync();

                    using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                    {
                        item.CopyTo(stream);
                    }

                    fileAttachment = new PreProcHeaderAttachment();

                    fileAttachment.FileName = filename;
                    fileAttachment.StrictedFileName = fileNameWithPath;
                    fileAttachment.Folder = preproc.PresalesId;
                    fileAttachment.PreProcId = Int32.Parse(preprocid);
                    fileAttachment.DocumentTypeId = tempdoctype.Id;
                    //if (detailid != 0)
                    //{
                    //    fileAttachment.Bastid = tempDetail.DetailId;
                    //}

                    _context.Add(fileAttachment);
                    await _context.SaveChangesAsync();

                    idattach += fileAttachment.Id + ";";
                }
            }
            return idattach;
        }

        [HttpPost]
        public async Task<string> SaveAttachVendorQuot(List<IFormFile> fileUploadvendorquot, string preprocid, int detailid)
        {
            getUserNow();
            var data = new PreProcDetail();
            PreProcHeaderAttachment fileAttachment;
            PreProcDetail tempDetail;
            string idattach = "";

            //if (detailid == 0)
            //{
            //    tempDetail = new PreProcDetail();

            //    tempDetail.Id = Int32.Parse(preprocid);
            //    tempDetail.Spec = "Attach";
            //    //tempDetail.SpecAttachId = fileAttachment.Id;
            //    _context.Add(tempDetail);
            //    await _context.SaveChangesAsync();

            //    data = tempDetail;
            //}
            //else
            //{
            //    tempDetail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid);

            //    if (tempDetail != null)
            //    {
            //        tempDetail.Spec = "Attach";
            //        //tempDetail.SpecAttachId = fileAttachment.Id;
            //        await _context.SaveChangesAsync();

            //        data = tempDetail;
            //    }
            //}

            foreach (var item in fileUploadvendorquot)
            {
                if (fileUploadvendorquot != null)
                {
                    var preprocdet = await _context.PreProcDetails.Where(x => x.Id == detailid).FirstOrDefaultAsync();
                    var preproc = await _context.PreProcGeneralInfo2s.Where(x => x.Id == Int32.Parse(preprocid)).FirstOrDefaultAsync();

                    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                    //create folder if not exist
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);


                    string filename = DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_" + item.FileName;

                    string fileNameWithPath = Path.Combine(path, filename);

                    var doctype = from t in _context.MsDocTypes
                                  select t;

                    var tempdoctype = await doctype.Where(t => t.DocType.Equals("VQU")).FirstOrDefaultAsync();

                    using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                    {
                        item.CopyTo(stream);
                    }

                    fileAttachment = new PreProcHeaderAttachment();

                    fileAttachment.FileName = filename;
                    fileAttachment.StrictedFileName = fileNameWithPath;
                    fileAttachment.Folder = preproc.PresalesId;
                    fileAttachment.PreProcId = Int32.Parse(preprocid);
                    fileAttachment.DocumentTypeId = tempdoctype.Id;
                    //if (detailid != 0)
                    //{
                    //    fileAttachment.Bastid = tempDetail.DetailId;
                    //}

                    _context.Add(fileAttachment);
                    await _context.SaveChangesAsync();

                    idattach += fileAttachment.Id + ";";
                }
            }
            return idattach;
        }

        [HttpPost]
        public async Task<string> SaveAttachSurDuk(List<IFormFile> fileUploadsurduk, string preprocid, int detailid)
        {
            getUserNow();
            var data = new PreProcDetail();
            PreProcHeaderAttachment fileAttachment;
            PreProcDetail tempDetail;
            string idattach = "";

            //if (detailid == 0)
            //{
            //    tempDetail = new PreProcDetail();
            //    tempDetail.Id = Int32.Parse(preprocid);
            //    tempDetail.SurDukCheck = true;
            //    tempDetail.SurDuk = "Attach";
            //    //tempDetail.SurdukAttachId = fileAttachment.Id;
            //    _context.Add(tempDetail);
            //    await _context.SaveChangesAsync();

            //    data = tempDetail;
            //}
            //else
            //{
            //    tempDetail = await _context.PreProcDetails.FirstOrDefaultAsync(x => x.DetailId == detailid);

            //    if (tempDetail != null)
            //    {
            //        tempDetail.SurDukCheck = true;
            //        tempDetail.SurDuk = "Attach";
            //        //tempDetail.SurdukAttachId = fileAttachment.Id;
            //        await _context.SaveChangesAsync();

            //        data = tempDetail;
            //    }

            //}
            if (fileUploadsurduk != null)
            {
                foreach (var item in fileUploadsurduk)
                {

                    var preprocdet = await _context.PreProcDetails.Where(x => x.Id == detailid).FirstOrDefaultAsync();
                    var preproc = await _context.PreProcGeneralInfo2s.Where(x => x.Id == Int32.Parse(preprocid)).FirstOrDefaultAsync();

                    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.PresalesId);

                    //create folder if not exist
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);


                    string filename = DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_" + item.FileName;

                    string fileNameWithPath = Path.Combine(path, filename);

                    var doctype = from t in _context.MsDocTypes
                                  select t;

                    var tempdoctype = await doctype.Where(t => t.DocType.Equals("SRD")).FirstOrDefaultAsync();

                    using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                    {
                        item.CopyTo(stream);
                    }
                    fileAttachment = new PreProcHeaderAttachment();

                    fileAttachment.FileName = filename;
                    fileAttachment.StrictedFileName = fileNameWithPath;
                    fileAttachment.Folder = preproc.PresalesId;
                    fileAttachment.PreProcId = Int32.Parse(preprocid);
                    fileAttachment.DocumentTypeId = tempdoctype.Id;

                    _context.Add(fileAttachment);
                    await _context.SaveChangesAsync();

                    idattach += fileAttachment.Id + ";";
                }
            }
            return idattach;
        }

        [HttpPost]
        public async Task<string> SaveAttachGRN(List<IFormFile> fileUploadgrn, string preprocid, int detailid)
        {
            getUserNow();
            var data = new PreProcFileAttachment();
            PreProcFileAttachment fileAttachment;
            //TempPreProcDetail tempDetail;

            if (fileUploadgrn != null)
            {
                foreach (var item in fileUploadgrn)
                {

                    var preprocdet = await _context.PreProcDetails.Where(x => x.Id == detailid).FirstOrDefaultAsync();
                    var preproc = await _context.PreProcGeneralInfo2s.Where(x => x.Id == Int32.Parse(preprocid)).FirstOrDefaultAsync();

                    string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + preproc.Pid);

                    //create folder if not exist
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);


                    string filename = DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_" + item.FileName;

                    string fileNameWithPath = Path.Combine(path, filename);

                    var doctype = from t in _context.MsDocTypes
                                  select t;

                    var tempdoctype = await doctype.Where(t => t.DocType.Equals("GRN")).FirstOrDefaultAsync();

                    using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                    {
                        item.CopyTo(stream);
                    }
                    fileAttachment = new PreProcFileAttachment();

                    fileAttachment.FileName = filename;
                    fileAttachment.StrictedFileName = fileNameWithPath;
                    fileAttachment.Folder = preproc.PresalesId;
                    fileAttachment.PreProcId = Int32.Parse(preprocid);
                    fileAttachment.DocumentTypeId = tempdoctype.Id;
                    fileAttachment.Bastid = detailid;
                    fileAttachment.GrnattachId = detailid;

                    _context.Add(fileAttachment);
                    await _context.SaveChangesAsync();

                    data = fileAttachment;
                }
            }
            return "Sukses";
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            getUserNow();
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public FileResult GetTemplate()
        {
            getUserNow();
            string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/Template/");

            string fileNameWithPath = Path.Combine(path, "Template PreProc Item.xlsx");

            byte[] FileBytes = System.IO.File.ReadAllBytes(fileNameWithPath);

            //"application/force-download" opsi lain
            return File(FileBytes, "application/octet-stream", "Template PreProc Item.xlsx");

        }

        public async Task<IActionResult> ShowRDL(string type, int id)
        {
            getUserNow();
            //DEFAULT URL LINK => VIEW PO DATA BY PID

            var urlapp = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Url Report");

            var urllink = urlapp.AlphaNumericValue + "/ReportServer/Pages/ReportViewer.aspx?%2fReports%2fPO+Data+by+PID&ID=" + id ?? "";

            if (type == "View PreProc Details")
            {
                urllink = urlapp.AlphaNumericValue + "/ReportServer/Pages/ReportViewer.aspx?%2fReports%2fView+PreProc+Details+Item&PreProcID=" + id ?? "";
            }
            else if (type == "Tracking Delivery Sheet")
            {
                urllink = urlapp.AlphaNumericValue + "/ReportServer/Pages/ReportViewer.aspx?%2fReports%2fTracking+Delivery+Sheet&ListItemID=" + id ?? "";
            }

            return Redirect(urllink);
        }
        
        public async Task<IActionResult> Index(int pageNumber, int pageSize, string pid)
        {
            getUserNow();

            var usernow = User.Identity.Name.Split("\\");

            if (pageNumber == 0 && pageSize == 0)
            {
                pageNumber = 1;
                pageSize = 25;
            }
            else if (pageNumber == 0)
            {
                pageNumber = 1;
            }

            // Connection string
            string connectionString = _configuration.GetConnectionString("PreProcConnection");
            string query = @"
                SELECT *
                FROM PreProcGeneralInfo2 where 
                    ([Created By] is not null and [Created By] like '%" + usernow[1] + @"%') or 
		            ([PIC Proc] is not null and [PIC Proc] like '%" + usernow[1] + @"%') or 
                    ([Project Support] is not null and [Project Support] like '%" + usernow[1] + @"%') or 
		            ([PM] is not null and [PM] like '%" + usernow[1] + @"%') or 
		            ([PC] is not null and [PC] like '%" + usernow[1] + @"%') or 
		            ([Prop. Mgr] is not null and [Prop. Mgr] like '%" + usernow[1] + @"%') or 
		            ([Tender Adm] is not null and [Tender Adm] like '%" + usernow[1] + @"%') or 
		            ([Modified By] is not null and [Modified By] like '%" + usernow[1] + @"%') or 
		            ([EmailCommitteeDefault] is not null and [EmailCommitteeDefault] like '%" + usernow[1] + @"%') or 
		            ([CVPM] is not null and [CVPM] like '%" + usernow[1] + @"%') or 
		            ([CVPC] is not null and [CVPC] like '%" + usernow[1] + @"%') or
		            (AM is not null and AM like '%" + usernow[1] + @"%')
                ORDER BY ID DESC
                OFFSET @PageSize * (@PageNumber - 1) ROWS
                FETCH NEXT @PageSize ROWS ONLY;";

            ViewBag.pid = "";

            // SQL query for pagination
            if (pid != null)
            {
                query = @"SELECT * FROM PreProcGeneralInfo2 where [Presales ID] = '" + pid + "'" +
                    "And ([Created By] is not null and [Created By] like '%" + usernow[1] + @"%') or 
		            ([PIC Proc] is not null and [PIC Proc] like '%" + usernow[1] + @"%') or 
                    ([Project Support] is not null and [Project Support] like '%" + usernow[1] + @"%') or 
		            ([PM] is not null and [PM] like '%" + usernow[1] + @"%') or 
		            ([PC] is not null and [PC] like '%" + usernow[1] + @"%') or 
		            ([Prop. Mgr] is not null and [Prop. Mgr] like '%" + usernow[1] + @"%') or 
		            ([Tender Adm] is not null and [Tender Adm] like '%" + usernow[1] + @"%') or 
		            ([Modified By] is not null and [Modified By] like '%" + usernow[1] + @"%') or 
		            ([EmailCommitteeDefault] is not null and [EmailCommitteeDefault] like '%" + usernow[1] + @"%') or 
		            ([CVPM] is not null and [CVPM] like '%" + usernow[1] + @"%') or 
		            ([CVPC] is not null and [CVPC] like '%" + usernow[1] + @"%') or
		            (AM is not null and AM like '%" + usernow[1] + @"%')" +
                    " ORDER BY ID DESC OFFSET @PageSize * (@PageNumber - 1) ROWS FETCH NEXT @PageSize ROWS ONLY;";
                ViewBag.pid = pid;
            }

            // Execute query
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@PageNumber", pageNumber);
                command.Parameters.AddWithValue("@PageSize", pageSize);

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                var employee = await _context.VwEmployeeBases.Where(x=>x.EmployeeEmail != null).ToListAsync();

                ViewBag.datasource = dataTable;
                //IEnumerable<PreProcGeneralInfo2> preprocs = dataTable.AsEnumerable().Select(row => new PreProcGeneralInfo2
                //{
                //    // Map properties from the DataRow to the DeliveryOrderTemplate object
                //    ListItemId = row.Field<int>("ListItemId"),
                //    PicProc = row.Field<string>("Pic Proc") != null ? employee.Where(x => x.EmployeeEmail.ToLower() == row.Field<string>("AM").ToLower()).Select(x => x.EmployeeName).FirstOrDefault() : "",
                //    PresalesId = row.Field<string>("Presales ID"),
                //    Id = row.Field<int>("ID"),
                //    //Am = row.Field<string>("AM"),
                //    Am = row.Field<string>("AM") != null ? employee.Where(x=>x.EmployeeEmail.ToLower() == row.Field<string>("AM").ToLower()).Select(x=>x.EmployeeName).FirstOrDefault() : "",
                //    Pid = row.Field<string>("PID"),
                //    Description = row.Field<string>("Description"),
                //    CustomerCode = row.Field<string>("CustomerCode"),
                //    ProjectSupport = row.Field<string>("Project Support"),
                //    Pm = row.Field<string>("PM") != null ? employee.Where(x => x.EmployeeEmail.ToLower() == row.Field<string>("PM").ToLower()).Select(x => x.EmployeeName).FirstOrDefault() : "",
                //    Stage = row.Field<string>("Stage"),
                //    PropMgr  = row.Field<string>("Prop. Mgr") != null ? employee.Where(x => x.EmployeeEmail.ToLower() == row.Field<string>("Prop. Mgr").ToLower()).Select(x => x.EmployeeName).FirstOrDefault() : "",
                //    TenderAdm  = row.Field<string>("Tender Adm") != null ? employee.Where(x => x.EmployeeEmail.ToLower() == row.Field<string>("Tender Adm").ToLower()).Select(x => x.EmployeeName).FirstOrDefault() : "",
                //    CreationDate = row.Field<DateTime?>("Creation Date") ?? DateTime.Parse("01-01-1001"),
                //    Pc  = row.Field<string>("PC") != null ? employee.Where(x => x.EmployeeEmail.ToLower() == row.Field<string>("PC").ToLower()).Select(x => x.EmployeeName).FirstOrDefault() : "",
                //    EmailCommitteeDefault  = row.Field<string>("EmailCommitteeDefault"),
                //    CreatedBy  = row.Field<string>("Created By") != null ? employee.Where(x => x.EmployeeEmail.ToLower() == row.Field<string>("Created By").ToLower()).Select(x => x.EmployeeName).FirstOrDefault() : "",
                //    // Map other properties as needed
                //});

                IEnumerable<PreProcGeneralInfo2> preprocs = dataTable.AsEnumerable().Select(row =>
                {
                    var newItem = new PreProcGeneralInfo2
                    {
                        // Map properties from the DataRow to the DeliveryOrderTemplate object
                        ListItemId = row.Field<int>("ListItemId"),
                        PresalesId = row.Field<string>("Presales ID"),
                        Id = row.Field<int>("ID"),
                        Am = row.Field<string>("AM") != null ? employee.FirstOrDefault(x => x.DomainName.ToLower() == row.Field<string>("AM").ToLower())?.EmployeeName : "",
                        Pid = row.Field<string>("PID"),
                        Description = row.Field<string>("Description"),
                        Customer = row.Field<string>("Customer"),
                        ProjectSupport = row.Field<string>("Project Support"),
                        Pm = row.Field<string>("PM") != null ? employee.FirstOrDefault(x => x.DomainName.ToLower() == row.Field<string>("PM").ToLower())?.EmployeeName : "",
                        Stage = row.Field<string>("Stage"),
                        PropMgr = row.Field<string>("Prop. Mgr") != null ? employee.FirstOrDefault(x => x.DomainName.ToLower() == row.Field<string>("Prop. Mgr").ToLower())?.EmployeeName : "",
                        TenderAdm = row.Field<string>("Tender Adm") != null ? employee.FirstOrDefault(x => x.DomainName.ToLower() == row.Field<string>("Tender Adm").ToLower())?.EmployeeName : "",
                        CreationDate = row.Field<DateTime?>("Creation Date") ?? DateTime.Parse("01-01-1001"),
                        Pc = row.Field<string>("PC") != null ? employee.FirstOrDefault(x => x.DomainName.ToLower() == row.Field<string>("PC").ToLower())?.EmployeeName : "",
                        CreatedBy = row.Field<string>("Created By") != null ? employee.FirstOrDefault(x => x.DomainName.ToLower() == row.Field<string>("Created By").ToLower())?.EmployeeName : ""
                    };

                    // Process PicProc separately
                    if (row.Field<string>("Pic Proc") != null)
                    {
                        var splitpic = row.Field<string>("Pic Proc").Split(';');
                        foreach (var item1 in splitpic)
                        {
                            if (!string.IsNullOrEmpty(item1))
                            {
                                var picname = employee.FirstOrDefault(x => x.EmployeeEmail.ToLower() == item1.ToLower())?.EmployeeName;
                                if (picname != null)
                                {
                                    newItem.PicProc += picname + ";";
                                }
                            }
                        }
                    }
                    if (row.Field<string>("EmailCommitteeDefault") != null)
                    {
                        var splitcomt = row.Field<string>("EmailCommitteeDefault").Split(';');
                        foreach (var item1 in splitcomt)
                        {
                            if (!string.IsNullOrEmpty(item1))
                            {
                                var comtname = employee.FirstOrDefault(x => x.EmployeeEmail.ToLower() == item1.ToLower())?.EmployeeName;
                                if (comtname != null)
                                {
                                    newItem.EmailCommitteeDefault += comtname;
                                }
                            }
                        }
                    }
                    return newItem;
                });


                // Calculate total pages
                int totalCount = GetTotalCount(1);
                int totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
                ViewBag.pagenumber = pageNumber;
                ViewBag.pagesize = pageSize;
                ViewBag.totalcount = totalPages;
                ViewBag.totaldata = preprocs.Count();

                return View(preprocs);
            }
        }

        private int GetTotalCount(int i)
        {
            getUserNow();
            int TotalCount = 0;
            if (i == 1)
            {
                TotalCount = _context.PreProcGeneralInfo2s.Count();
            }
            
            return TotalCount;
        }

        [HttpPost]
        public async Task<PreProcPmochecklist> SendNotifToPMO(int preprocid)
        {
            getUserNow();
            var data = new PreProcPmochecklist();

            string TextClosed = @"<br><br><p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>&nbsp; <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US'>
                                            <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>Please Do Not Reply to this email because we are not monitoring this inbox<o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <br>
                                        <div class='elementToProof'>
                                          <div style='font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);'>
                                            <br>
                                          </div>
                                          <div id='Signature'>
                                            <div>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <span style='font-size:8pt;font-family:Times New Roman,serif'>Best Regards,&nbsp;</span>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='font-family:Times New Roman,serif'>Internal System&nbsp;</span>
                                                </b>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'>PT MULTIPOLAR TECHNO</span>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> LOGY Tbk(MLPT)</span>
                                                </b>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> &nbsp;</span>
                                                </b>
                                              </p>
                                            </div>
                                          </div>
                                        </div>
                                        </body>
                                    </html>";

            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "10.10.62.17";
                builder.UserID = "sa";
                builder.Password = "Password1!";
                builder.InitialCatalog = "eRequisition";
                builder.TrustServerCertificate = true;
                builder.MultipleActiveResultSets = true;

                var paramemail = await _context.MsParameterValues.Where(x=>x.Parameter == "Parameter Email Notif To PMO" && x.Title == "PreProc").ToListAsync();

                var proc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == preprocid);
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    var mailMessage = new MailMessage();
                    mailMessage.Bcc.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("adam.takariyanto@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("internal.apps@multipolar.com"));
                    mailMessage.From = new MailAddress("mlpt365@multipolar.com", "Multipolar Technology 365");
                    mailMessage.Subject = paramemail[0].AlphaNumericValue + " " + proc.Pid;
                    mailMessage.Body = paramemail[1].AlphaNumericValue;
                    mailMessage.IsBodyHtml = true;

                    string CurrentGroup = "";
                    string Status = "";
                    string Level2 = "";
                    string Level4 = "";
                    string Level5 = "";
                    int Level = 0;

                    String sql = "SELECT * FROM gp_ms_Emp WHERE EmailAddress = '" + proc.Pm + "'";

                    var urlapp = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Url App");
                    
                    var additionrecpt = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT");
                    var additionrecptpmo = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT PMO").ToListAsync();

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                mailMessage.Body += urlapp.AlphaNumericValue + "/Preprocs/ViewDetail/" + proc.Id + "'>Click here to open this item</a> ";
                                //mailMessage.To.Add(new MailAddress(reader[18].ToString()));
                                Status = reader[12].ToString();
                                Level2 = reader[4].ToString();
                                Level4 = reader[7].ToString();
                                Level5 = reader[9].ToString();
                                Level = Convert.ToInt32(reader[13].ToString());
                                if (Status == "m")
                                {
                                    if (Level == 6)
                                    {
                                        CurrentGroup = Level5;
                                    }
                                    else if (Level == 5)
                                    {
                                        CurrentGroup = Level4;
                                    }
                                    else if (Level == 4)
                                    {
                                        CurrentGroup = Level2;
                                    }
                                }
                                else
                                {
                                    CurrentGroup = reader[14].ToString();
                                }
                                mailMessage.To.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));
                                if (additionrecpt != null)
                                {
                                    mailMessage.To.Add(new MailAddress(additionrecpt.AlphaNumericValue));

                                }
                                if (additionrecptpmo != null){
                                    foreach (var addrecptpmo in additionrecptpmo)
                                    {
                                        mailMessage.To.Add(new MailAddress(addrecptpmo.AlphaNumericValue));
                                    }
                                }
                            }
                        }
                    }

                    connection.Close();

                    mailMessage.Body += TextClosed;

                    try
                    {
                        using (var smtpClient = new SmtpClient())
                        {
                            smtpClient.Host = "10.10.62.18"; //production 64.229
                            smtpClient.Send(mailMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (SqlException e)
            {
                //break;
            }

            //return Ok();
            return data;
        }

        [HttpPost]
        public async Task<PreProcPendingBillingReason> SendNotifToPMPC(string message, int procid)
        {
            getUserNow();
            var data = new PreProcPendingBillingReason();

            string TextOpen = @"<html>
                                      <head>
                                         <style>
                                        <!--
                                         /* Font Definitions */
                                         @font-face
	                                        {font-family:""Cambria Math"";
	                                        panose-1:2 4 5 3 5 4 6 3 2 4;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:roman;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:3 0 0 0 1 0;}
                                        @font-face
	                                        {font-family:Calibri;
	                                        panose-1:2 15 5 2 2 2 4 3 2 4;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:swiss;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:-469750017 -1073732485 9 0 511 0;}
                                        @font-face
	                                        {font-family:""Segoe UI"";
	                                        panose-1:2 11 5 2 4 2 4 2 2 3;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:swiss;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:-469750017 -1073683329 9 0 511 0;}
                                         /* Style Definitions */
                                         p.MsoNormal, li.MsoNormal, div.MsoNormal
	                                        {mso-style-unhide:no;
	                                        mso-style-qformat:yes;
	                                        mso-style-parent:"""";
	                                        margin:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        a:link, span.MsoHyperlink
	                                        {mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        color:blue;
	                                        text-decoration:underline;
	                                        text-underline:single;}
                                        a:visited, span.MsoHyperlinkFollowed
	                                        {mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        color:purple;
	                                        text-decoration:underline;
	                                        text-underline:single;}
                                        p
	                                        {mso-style-priority:99;
	                                        mso-margin-top-alt:auto;
	                                        margin-right:0in;
	                                        mso-margin-bottom-alt:auto;
	                                        margin-left:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        p.msonormal0, li.msonormal0, div.msonormal0
	                                        {mso-style-name:msonormal;
	                                        mso-style-unhide:no;
	                                        mso-margin-top-alt:auto;
	                                        margin-right:0in;
	                                        mso-margin-bottom-alt:auto;
	                                        margin-left:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        .MsoChpDefault
	                                        {mso-style-type:export-only;
	                                        mso-default-props:yes;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-ascii-font-family:Calibri;
	                                        mso-ascii-theme-font:minor-latin;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;
	                                        mso-hansi-font-family:Calibri;
	                                        mso-hansi-theme-font:minor-latin;
	                                        mso-bidi-font-family:""Times New Roman"";
	                                        mso-bidi-theme-font:minor-bidi;}
                                        @page WordSection1
	                                        {size:8.5in 11.0in;
	                                        margin:1.0in 1.0in 1.0in 1.0in;
	                                        mso-header-margin:.5in;
	                                        mso-footer-margin:.5in;
	                                        mso-paper-source:0;}
                                        div.WordSection1
	                                        {page:WordSection1;}
                                         table.MsoNormalTable
	                                        {mso-style-name:""Table Normal"";
	                                        mso-tstyle-rowband-size:0;
	                                        mso-tstyle-colband-size:0;
	                                        mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        mso-style-parent:"""";
	                                        mso-padding-alt:0in 5.4pt 0in 5.4pt;
	                                        mso-para-margin:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-ascii-font-family:Calibri;
	                                        mso-ascii-theme-font:minor-latin;
	                                        mso-hansi-font-family:Calibri;
	                                        mso-hansi-theme-font:minor-latin;
	                                        mso-bidi-font-family:""Times New Roman"";
	                                        mso-bidi-theme-font:minor-bidi;}
                                        </style>
                                       </head>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>Dear ";

            string TextOpenAfterHead = @",</span>
                                        </p>";

            string TextClosed = @"<br><br><p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>&nbsp; <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US'>
                                            <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>Please Do Not Reply to this email because we are not monitoring this inbox<o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <br>
                                        <div class='elementToProof'>
                                          <div style='font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);'>
                                            <br>
                                          </div>
                                          <div id='Signature'>
                                            <div>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <span style='font-size:8pt;font-family:Times New Roman,serif'>Best Regards,&nbsp;</span>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='font-family:Times New Roman,serif'>Internal System&nbsp;</span>
                                                </b>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'>PT MULTIPOLAR TECHNO</span>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> LOGY Tbk(MLPT)</span>
                                                </b>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> &nbsp;</span>
                                                </b>
                                              </p>
                                            </div>
                                          </div>
                                        </div>
                                        </body>
                                    </html>";

            string TdOpener = @"<td style='border:none;padding:3.0pt 3.0pt 3.0pt 3.0pt'><p class=MsoNormal><span style='font-family:""Segoe UI"",sans-serif'>";
            string TdCloser = @"<o:p></o:p></span></p></td>";
            //------------------------------------------------------------------------------------------------------------------------ REMINDER DELIVERY ORDER BORROW ---------------------------------------------------------------------------------------
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "10.10.62.17";
                builder.UserID = "sa";
                builder.Password = "Password1!";
                builder.InitialCatalog = "eRequisition";
                builder.TrustServerCertificate = true;
                builder.MultipleActiveResultSets = true;
                var proc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    var mailMessage = new MailMessage();
                    mailMessage.Bcc.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("adam.takariyanto@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("internal.apps@multipolar.com"));
                    mailMessage.From = new MailAddress("mlpt365@multipolar.com", "Multipolar Technology 365");
                    mailMessage.Subject = "[PREPROC] Assignment Member Project ID " + proc.Pid;
                    mailMessage.Body = TextOpen;
                    mailMessage.IsBodyHtml = true;
                    string TableHeader = @"<table class=MsoNormalTable border=1 cellspacing=3 cellpadding=0 style='mso-cellspacing:1.5pt;border:solid #104E8B 1.0pt;mso-border-alt:solid #104E8B .75pt; mso-yfti-tbllook:1184;mso-padding-alt:3.0pt 3.0pt 3.0pt 3.0pt'>
                                        <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>DO Number<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Borrower Name<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Due Date<o:p></o:p></span></b></p>
                                            </td>
                                            <td style='border:none;background:#104E8B;padding:3.0pt 3.0pt 3.0pt 3.0pt'>
                                                <p class=MsoNormal><b><span style='font-family:""Segoe UI"",sans-serif;
                                      color:white'>Pending Days (Before Due Date)<o:p></o:p></span></b></p>
                                            </td></tr>";

                    string CurrentGroup = "";
                    string Status = "";
                    string Level2 = "";
                    string Level4 = "";
                    string Level5 = "";
                    int Level = 0;

                    String sql = "SELECT * FROM gp_ms_Emp WHERE EmailAddress = '" + proc.Pm + "'";

                    var urlapp = await _context.MsParameterValues.FirstOrDefaultAsync(x=>x.Title == "PreProc" && x.Parameter == "Url App");

                    var additionrecpt = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT");
                    var additionrecptpmo = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT PMO").ToListAsync();

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                mailMessage.Body += reader[1].ToString() + TextOpenAfterHead;
                                mailMessage.Body += "<br>" + message + "<br><br> You have been assigned to be PM for this item. <br> Please click the hyperlink below for further details: <br><br> <a  class='btn-down-spec' style='font-size: 12px; padding: 7px!important;margin-right: 5px;margin-left: 5px;' href='" + urlapp.AlphaNumericValue + "/Preprocs/ViewDetail/" + proc.Id + "'>Click here to open this item</a> ";
                                //mailMessage.To.Add(new MailAddress(reader[18].ToString()));
                                Status = reader[12].ToString();
                                Level2 = reader[4].ToString();
                                Level4 = reader[7].ToString();
                                Level5 = reader[9].ToString();
                                Level = Convert.ToInt32(reader[13].ToString());
                                if (Status == "m")
                                {
                                    if (Level == 6)
                                    {
                                        CurrentGroup = Level5;
                                    }
                                    else if (Level == 5)
                                    {
                                        CurrentGroup = Level4;
                                    }
                                    else if (Level == 4)
                                    {
                                        CurrentGroup = Level2;
                                    }
                                }
                                else
                                {
                                    CurrentGroup = reader[14].ToString();
                                }
                                mailMessage.To.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));

                                if (additionrecpt != null)
                                {
                                    mailMessage.To.Add(new MailAddress(additionrecpt.AlphaNumericValue));

                                }
                                if (additionrecptpmo != null) {
                                    foreach (var addrecptpmo in additionrecptpmo) {
                                        mailMessage.To.Add(new MailAddress(addrecptpmo.AlphaNumericValue));
                                    }
                                }
                            }
                        }
                    }

                    connection.Close();

                    mailMessage.Body += TextClosed;

                    try
                    {
                        using (var smtpClient = new SmtpClient())
                        {
                            smtpClient.Host = "10.10.62.18"; //production 64.229
                            smtpClient.Send(mailMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (SqlException e)
            {
                //break;
            }

            //return Ok();
            return data;
        }


        [HttpPost]
        public async Task<PreProcPendingBillingReason> SendNotifStagePresales(string message, int procid, string type, string recipient)
        {
            getUserNow();
            var data = new PreProcPendingBillingReason();

            string TextOpen = @"<html>
                                      <head>
                                         <style>
                                        <!--
                                         /* Font Definitions */
                                         @font-face
	                                        {font-family:""Cambria Math"";
	                                        panose-1:2 4 5 3 5 4 6 3 2 4;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:roman;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:3 0 0 0 1 0;}
                                        @font-face
	                                        {font-family:Calibri;
	                                        panose-1:2 15 5 2 2 2 4 3 2 4;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:swiss;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:-469750017 -1073732485 9 0 511 0;}
                                        @font-face
	                                        {font-family:""Segoe UI"";
	                                        panose-1:2 11 5 2 4 2 4 2 2 3;
	                                        mso-font-charset:0;
	                                        mso-generic-font-family:swiss;
	                                        mso-font-pitch:variable;
	                                        mso-font-signature:-469750017 -1073683329 9 0 511 0;}
                                         /* Style Definitions */
                                         p.MsoNormal, li.MsoNormal, div.MsoNormal
	                                        {mso-style-unhide:no;
	                                        mso-style-qformat:yes;
	                                        mso-style-parent:"""";
	                                        margin:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        a:link, span.MsoHyperlink
	                                        {mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        color:blue;
	                                        text-decoration:underline;
	                                        text-underline:single;}
                                        a:visited, span.MsoHyperlinkFollowed
	                                        {mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        color:purple;
	                                        text-decoration:underline;
	                                        text-underline:single;}
                                        p
	                                        {mso-style-priority:99;
	                                        mso-margin-top-alt:auto;
	                                        margin-right:0in;
	                                        mso-margin-bottom-alt:auto;
	                                        margin-left:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        p.msonormal0, li.msonormal0, div.msonormal0
	                                        {mso-style-name:msonormal;
	                                        mso-style-unhide:no;
	                                        mso-margin-top-alt:auto;
	                                        margin-right:0in;
	                                        mso-margin-bottom-alt:auto;
	                                        margin-left:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;}
                                        .MsoChpDefault
	                                        {mso-style-type:export-only;
	                                        mso-default-props:yes;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-ascii-font-family:Calibri;
	                                        mso-ascii-theme-font:minor-latin;
	                                        mso-fareast-font-family:Calibri;
	                                        mso-fareast-theme-font:minor-latin;
	                                        mso-hansi-font-family:Calibri;
	                                        mso-hansi-theme-font:minor-latin;
	                                        mso-bidi-font-family:""Times New Roman"";
	                                        mso-bidi-theme-font:minor-bidi;}
                                        @page WordSection1
	                                        {size:8.5in 11.0in;
	                                        margin:1.0in 1.0in 1.0in 1.0in;
	                                        mso-header-margin:.5in;
	                                        mso-footer-margin:.5in;
	                                        mso-paper-source:0;}
                                        div.WordSection1
	                                        {page:WordSection1;}
                                         table.MsoNormalTable
	                                        {mso-style-name:""Table Normal"";
	                                        mso-tstyle-rowband-size:0;
	                                        mso-tstyle-colband-size:0;
	                                        mso-style-noshow:yes;
	                                        mso-style-priority:99;
	                                        mso-style-parent:"""";
	                                        mso-padding-alt:0in 5.4pt 0in 5.4pt;
	                                        mso-para-margin:0in;
	                                        mso-pagination:widow-orphan;
	                                        font-size:11.0pt;
	                                        font-family:""Calibri"",sans-serif;
	                                        mso-ascii-font-family:Calibri;
	                                        mso-ascii-theme-font:minor-latin;
	                                        mso-hansi-font-family:Calibri;
	                                        mso-hansi-theme-font:minor-latin;
	                                        mso-bidi-font-family:""Times New Roman"";
	                                        mso-bidi-theme-font:minor-bidi;}
                                        </style>
                                       </head>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>";

            
            string TextClosed = @"<br><br><p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>&nbsp; <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US'>
                                            <o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <p class='xmsonormal' style='margin:0cm 0cm 0.0001pt;font-size:11pt;font-family:Calibri, sans-serif'>
                                          <span lang='EN-US' style='mso-ansi-language:EN-US' class='ContentPasted0'>Please Do Not Reply to this email because we are not monitoring this inbox<o:p class='ContentPasted0'>&nbsp;</o:p>
                                          </span>
                                        </p>
                                        <br>
                                        <div class='elementToProof'>
                                          <div style='font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);'>
                                            <br>
                                          </div>
                                          <div id='Signature'>
                                            <div>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <span style='font-size:8pt;font-family:Times New Roman,serif'>Best Regards,&nbsp;</span>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='font-family:Times New Roman,serif'>Internal System&nbsp;</span>
                                                </b>
                                              </p>
                                              <p style='font-size:11pt;font-family:Calibri,sans-serif;margin:0'>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'>PT MULTIPOLAR TECHNO</span>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> LOGY Tbk(MLPT)</span>
                                                </b>
                                                <b>
                                                  <span style='color: rgb(64, 64, 64); font-size: 9pt; font-family: Times New Roman, serif;'> &nbsp;</span>
                                                </b>
                                              </p>
                                            </div>
                                          </div>
                                        </div>
                                        </body>
                                    </html>";

            //------------------------------------------------------------------------------------------------------------------------ REMINDER DELIVERY ORDER BORROW ---------------------------------------------------------------------------------------
            try
            {
                SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();

                builder.DataSource = "10.10.62.17";
                builder.UserID = "sa";
                builder.Password = "Password1!";
                builder.InitialCatalog = "eRequisition";
                builder.TrustServerCertificate = true;
                builder.MultipleActiveResultSets = true;
                var proc = await _context.PreProcGeneralInfo2s.FirstOrDefaultAsync(x => x.Id == procid);
                using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
                {
                    connection.Open();

                    var mailMessage = new MailMessage();
                    mailMessage.Bcc.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("adam.takariyanto@multipolar.com"));
                    //mailMessage.Bcc.Add(new MailAddress("internal.apps@multipolar.com"));
                    mailMessage.From = new MailAddress("mlpt365@multipolar.com", "Multipolar Technology 365");
                    
                    mailMessage.Body = TextOpen;
                    mailMessage.IsBodyHtml = true;

                    string CurrentGroup = "";
                    string Status = "";
                    string Level2 = "";
                    string Level4 = "";
                    string Level5 = "";
                    int Level = 0;

                    String sql = "SELECT * FROM gp_ms_Emp WHERE EmailAddress = '" + proc.Pm + "'";

                    var additionrecpt = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT");
                    var additionrecptpmo = await _context.MsParameterValues.Where(x => x.Title == "PreProc" && x.Parameter == "Recipient for UAT PMO").ToListAsync();

                    using (SqlCommand command = new SqlCommand(sql, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                
                                //mailMessage.To.Add(new MailAddress(reader[18].ToString()));
                                Status = reader[12].ToString();
                                Level2 = reader[4].ToString();
                                Level4 = reader[7].ToString();
                                Level5 = reader[9].ToString();
                                Level = Convert.ToInt32(reader[13].ToString());
                                if (Status == "m")
                                {
                                    if (Level == 6)
                                    {
                                        CurrentGroup = Level5;
                                    }
                                    else if (Level == 5)
                                    {
                                        CurrentGroup = Level4;
                                    }
                                    else if (Level == 4)
                                    {
                                        CurrentGroup = Level2;
                                    }
                                }
                                else
                                {
                                    CurrentGroup = reader[14].ToString();
                                }
                                mailMessage.To.Add(new MailAddress("stepanus.triatmaja@multipolar.com"));

                                if (additionrecpt != null)
                                {
                                    mailMessage.To.Add(new MailAddress(additionrecpt.AlphaNumericValue));

                                }
                            }
                        }
                    }

                    connection.Close();

                    var parambody = await _context.MsParameterValues.Where(x=> x.Parameter.Equals(type)).OrderBy(x=>x.Level).ToListAsync();

                    mailMessage.Subject = parambody[0].AlphaNumericValue + proc.Pid;
                    mailMessage.Body += parambody[1].AlphaNumericValue;

                    var urlapp = await _context.MsParameterValues.FirstOrDefaultAsync(x => x.Title == "PreProc" && x.Parameter == "Url App");

                    if (type == "Parameter Email Excpt To Finance")
                    {
                        mailMessage.Body += message + "<br> Please click the hyperlink below for further details: <br><br> <a class='btn-down-spec' style='font-size: 12px; padding: 7px!important;margin-right: 5px;margin-left: 5px;' href='" + urlapp.AlphaNumericValue + "/Preprocs/ViewDetail/" + procid;
                        //mailMessage.To.Add(recipient);
                    }

                    mailMessage.Body += proc.Id + "'>Click here to open this item</a> ";

                    mailMessage.Body += TextClosed;

                    try
                    {
                        using (var smtpClient = new SmtpClient())
                        {
                            smtpClient.Host = "10.10.62.18"; //production 64.229
                            smtpClient.Send(mailMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (SqlException e)
            {
                //break;
            }

            //return Ok();
            return data;
        }

        [HttpPost]
        public async Task<PreProcDetail> CheckExistingItemDetailToOverwrite(string procid)
        {
            getUserNow();
            var data = await _context.PreProcDetails.FirstOrDefaultAsync(x=>x.Id == Int32.Parse(procid));

            return data;
        }

        [HttpPost]
        public async Task<PreProcHeaderAttachment> UploadAttachment(IFormFile fileUpload, string presalesid, string type)
        {
            var user = getUserNow();
            var fileAttachment = new PreProcHeaderAttachment();
            var doctype = new MsDocType();

            if (type == "RFP")
            {
                doctype = await _context.MsDocTypes.Where(x => x.DocType == "RFP").FirstOrDefaultAsync();
            }

            var data = new PreProcGeneralInfo2();
            data.PresalesId = presalesid;
            await _context.SaveChangesAsync();

            if (fileUpload != null)
            {
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/Files/" + data.PresalesId);

                //create folder if not exist
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);


                string filename = DateTime.Now.ToString("dd-MM-yyyy_HH_mm_ss") + "_" + fileUpload.FileName;

                string fileNameWithPath = Path.Combine(path, filename);

                using (var stream = new FileStream(fileNameWithPath, FileMode.Create))
                {
                    fileUpload.CopyTo(stream);
                }

                fileAttachment.FileName = filename;
                fileAttachment.StrictedFileName = fileNameWithPath;
                fileAttachment.Folder = data.PresalesId;
                fileAttachment.PreProcId = data.Id;
                fileAttachment.DocumentTypeId = doctype.Id;
                fileAttachment.HasBeenUploaded = true;
                fileAttachment.UploadedFrom = "PREPROC";
                fileAttachment.ModifiedBy = user;

                _context.Add(fileAttachment);
                await _context.SaveChangesAsync();
            }
            return fileAttachment;
        }

        public async Task<List<PreProcDetail>> GetItemDetailFromSearch(int procid, string searchpid)
        {
            var itemdet = await _context.PreProcDetails.Where(x => x.Id == procid && x.Idnew != "EDIT").ToListAsync();

            if (searchpid != null)
            {
                ViewBag.itemdetail = await _context.PreProcDetails.Where(x => x.Id == procid && x.Idnew != "EDIT" && x.Pidfull == searchpid).OrderBy(x => x.DetailId).ToListAsync();
            }

            return itemdet;
        }

        [HttpPost]
        public async Task<VwGetAllPid> GetDetailDataPid(string pid)
        {
            var data = new VwGetAllPid();

            if (pid != null)
            {
                var tempdata = await _context.VwGetAllPids.FirstOrDefaultAsync(x => x.ProjectId == pid);

                if (tempdata != null)
                {
                    data = tempdata;
                }
            }

            return data;
        }

        [HttpPost]
        public async Task<VwGetAllOptyId> GetDetailDataPresalesId(string presalesid)
        {
            var data = new VwGetAllOptyId();

            if (presalesid != null)
            {
                var tempdata = await _context.VwGetAllOptyIds.FirstOrDefaultAsync(x => x.NewPresalesid == presalesid);

                if (tempdata != null)
                {
                    data = tempdata;
                }
            }

            return data;
        }

        private bool PreProcExists(int id)
        {
            getUserNow();
            return (_context.PreProcGeneralInfo2s?.Any(e => e.Id == id)).GetValueOrDefault();
        }
    }
}
