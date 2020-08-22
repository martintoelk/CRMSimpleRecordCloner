using System;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;
using XrmToolBox.Extensibility.Args;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System.ComponentModel;
using System.Text;
using System.Collections.Specialized;
using Microsoft.Xrm.Tooling.Connector;

namespace martintmg.MSDYN.Tools.SimpleRecordCloner
{
    public partial class PluginControl : MultipleConnectionsPluginControlBase, IXrmToolBoxPluginControl, IGitHubPlugin, IStatusBarMessenger, IHelpPlugin
    {
        private Panel infoPanel;

        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;

        public string RepositoryName
        {
            get
            {
                return "CRMSimpleRecordCloner";
            }
        }

        public string UserName
        {
            get
            {
                return "martintmg";
            }
        }

        public string HelpUrl
        {
            get
            {
                return "https://github.com/martintmg/CRMSimpleRecordCloner";
            }
        }

        private void SetConnectionLabel(ConnectionDetail detail, string serviceType)
        {
            switch (serviceType)
            {
                case "Source":
                    lblSource.Text = detail.ConnectionName;
                    lblSource.ForeColor = Color.Green;
                    break;
            }
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName = "", object parameter = null)
        {
            ConnectionDetail = detail;
            if (actionName == "AdditionalOrganization")
            {
                AdditionalConnectionDetails.Clear();
                AdditionalConnectionDetails.Add(detail);
                SetConnectionLabel(detail, "Target");
            }
            else
            {
                SetConnectionLabel(detail, "Source");
            }

            base.UpdateConnection(newService, detail, actionName, parameter);
        }

        #region Base tool implementation

        public PluginControl()
        {
            InitializeComponent();
        }


        #endregion Base tool implementation

        private void btnChooseTarget_Click(object sender, EventArgs e)
        {
            AddAdditionalOrganization();
        }
        private void chkIgnoreAllLookups_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == true)
            {
                chkIgnoreOwnerAndModifiedBy.Checked = true;
                chkVerifyLookups.Checked = false;
            }
        }

        private void chkIgnoreOwnerAndModifiedBy_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == false)
                chkIgnoreAllLookups.Checked = false;
        }

        private void lstTargetEnvironments_KeyDown(object sender, KeyEventArgs e)
        {
            if (listBoxTargetEnv.SelectedItems.Count > 0 && e.KeyCode == Keys.Delete)
            {
                var conn = listBoxTargetEnv.SelectedItem as ConnectionDetail;

                if (conn != null)
                {
                    RemoveAdditionalOrganization(conn);
                }
            }
        }


        private void ToggleWaitMode(bool on)
        {
            if (on)
            {
                Cursor = Cursors.WaitCursor;
                btnAddToRecordList.Enabled = false;
                btnChooseTarget.Enabled = false;
                btnCloneRecord.Enabled = false;
                txtRecordURL.Enabled = false;
                chkIgnoreAllLookups.Enabled = false;
                chkIgnoreOwnerAndModifiedBy.Enabled = false;
                chkVerifyLookups.Enabled = false;
            }
            else
            {

            }
        }

        private void WorkerProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            InformationPanel.ChangeInformationPanelMessage(infoPanel, e.UserState.ToString());
        }

        private void WorkerRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            infoPanel.Dispose();
            Controls.Remove(infoPanel);

            btnAddToRecordList.Enabled = true;
            btnChooseTarget.Enabled = true;
            btnCloneRecord.Enabled = true;
            txtRecordURL.Enabled = true;
            chkIgnoreAllLookups.Enabled = true;
            chkIgnoreOwnerAndModifiedBy.Enabled = true;
            chkVerifyLookups.Enabled = true;

            Cursor = Cursors.Default;

            string message;

            if (e.Error != null)
            {
                message = string.Format("An error occured: {0}", e.Error.Message);
                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                message = "Finished Processing!";
                MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void WorkerDoWorkCloneRecords(object sender, DoWorkEventArgs e)
        {
            var bw = (BackgroundWorker)sender;
            var current = 0;

            foreach (var connectioNDetail in AdditionalConnectionDetails)
            {
                var targetService = connectioNDetail.GetCrmServiceClient();

                RetrieveAllEntitiesResponse SourceMetaData = GetMetaData();
                //  lastTargetService = targetService;

                var recordUrlsToProcess = (ListBox.ObjectCollection)e.Argument;

                foreach (var recordUrl in recordUrlsToProcess)
                {
                    current++;
                    var paramaters = GetParameterFromURL(recordUrl.ToString());

                    var RecordId = new Guid(GetParameter(paramaters, "id").ToString());
                    var logicalName = GetParameter(paramaters, "etn")?.ToString();
                    if (logicalName == null)
                    {
                        // fallback if the URL was generated through an oder Dynamics version
                        int objectTypeCode = int.Parse(GetParameter(paramaters, "etc").ToString());
                        logicalName = GetEntityLogicalNameFromMetadataByObjectTypeCode(objectTypeCode);
                    }

                    var entity = Service.Retrieve(logicalName, RecordId, new ColumnSet(true));

                    string displayName = GetRecordDisplayName(SourceMetaData, logicalName, entity);

                    bw.ReportProgress(0, string.Format("Cloning Record {4} of {5}: Displayname \"{0}\" LogicalName \"{1}\" Id \"{2}\" to Environemt \"{3}\"",
                        displayName, entity.LogicalName, entity.Id, targetService.ConnectedOrgFriendlyName,
                        current,
                        recordUrlsToProcess.Count * AdditionalConnectionDetails.Count()));

                    ApplyCloneRulesAndCloneRecord(targetService, SourceMetaData, logicalName, entity);
                }
            }
        }

        private void ApplyCloneRulesAndCloneRecord(CrmServiceClient TargetOrgSvc, RetrieveAllEntitiesResponse SourceMetaData, string logicalName, Entity RecordToClone)
        {
            RemoveTraversedPathAttributeFromEntity(RecordToClone);
            ApplyRemoveAllERefsRule(RecordToClone);
            ApplyIngoreOnwerModfiedByRule(RecordToClone);
            ApplyIgnoreStatusCodeRule(RecordToClone);
            ApplyRemoveNotFoundLookupsRule(TargetOrgSvc, RecordToClone);
            ApplyRemoveNonExistingAttributesInTargetOrganizationRule(TargetOrgSvc, RecordToClone);

            if (!ApplyAreAllLookupsPresentRule(TargetOrgSvc, RecordToClone))
                return;
            ApplyUpsertRecords(TargetOrgSvc, SourceMetaData, logicalName, RecordToClone);
        }

        private void RemoveAttributesFromRecord(List<string> Attributes, Entity RecordToClone)
        {
            Attributes.ForEach(attribute =>
            {
                if (RecordToClone.Contains(attribute))
                    RecordToClone.Attributes.Remove(attribute);
            });
        }

        private void RemoveTraversedPathAttributeFromEntity(Entity RecordToClone)
        {
            RemoveAttributesFromRecord(new List<string> { "traversedpath" }, RecordToClone);
        }

        private void ApplyRemoveAllERefsRule(Entity RecordToClone)
        {
            if (chkIgnoreAllLookups.Checked)
            {
                RemoveAttributesFromRecord(
                    GetAllERefsFromRecord(RecordToClone)
                    .Select(attribute => attribute.Key).ToList()
                    , RecordToClone);
            }
        }

        private static IEnumerable<KeyValuePair<string, object>> GetAllERefsFromRecord(Entity RecordToClone)
        {
            return RecordToClone.Attributes.Where(attribute => attribute.Value.GetType().Name == "EntityReference");
        }

        private void ApplyIngoreOnwerModfiedByRule(Entity RecordToClone)
        {
            if (chkIgnoreOwnerAndModifiedBy.Checked)
            {
                RemoveAttributesFromRecord(new List<string> { Helpers.EntityNames.Owner, Helpers.EntityNames.ModifiedBy }, RecordToClone);
            }
        }

        private void ApplyIgnoreStatusCodeRule(Entity RecordToClone)
        {
            if (chkIgnoreStateCode.Checked)
            {
                RemoveAttributesFromRecord(new List<string> { Helpers.EntityNames.StateCode, Helpers.EntityNames.StatusCode }, RecordToClone);
            }
        }

        private void ApplyRemoveNotFoundLookupsRule(IOrganizationService TargetOrgSvc, Entity RecordToClone)
        {
            if (chkRemoveNotFoundLookups.Checked)
            {
                List<EntityReference> notFoundERefs = GetERefsWhichNotExistsInTargetEnvironment(TargetOrgSvc, RecordToClone);

                RemoveAttributesFromRecord(notFoundERefs.Select(x => x.Name).ToList(), RecordToClone);
            }
        }

        private static List<EntityReference> GetERefsWhichNotExistsInTargetEnvironment(IOrganizationService TargetOrgSvc, Entity RecordToClone)
        {
            var allERefsAttributes = GetAllERefsFromRecord(RecordToClone).ToList();

            var notFoundERefs = new List<EntityReference>();

            allERefsAttributes.ForEach(eRefAttribute =>
            {

                EntityReference eRef = new EntityReference();
                try
                {
                    eRef = (EntityReference)eRefAttribute.Value;
                    var entity = TargetOrgSvc.Retrieve(eRef.LogicalName, eRef.Id, new ColumnSet(false));
                }
                catch
                {
                    notFoundERefs.Add(eRef);
                }
            });
            return notFoundERefs;
        }

        private void ApplyRemoveNonExistingAttributesInTargetOrganizationRule(CrmServiceClient TragetOrgService, Entity RecordToClone)
        {
            if (chkRemoveNonExistingAttributes.Checked)
            {
                RemoveMissingAttributeInTargetEnvironment(TragetOrgService, RecordToClone);
            }
        }

        private void RemoveMissingAttributeInTargetEnvironment(CrmServiceClient TragetOrgService, Entity RecordToClone)
        {
            var entityMetaData = new RetrieveEntityRequest
            {
                RetrieveAsIfPublished = true,
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                LogicalName = RecordToClone.LogicalName
            };

            var SourceMetaData = (RetrieveEntityResponse)TragetOrgService.Execute(entityMetaData);

            var sourceAttributes = RecordToClone.Attributes;

            var tempEntity = new Entity();

            SourceMetaData.EntityMetadata.Attributes.ToList().ForEach(attributeMetadata =>
            {
                var attributeLogicalName = attributeMetadata.LogicalName;

                if (!attributeMetadata.IsValidForCreate.Value)
                {
                    sourceAttributes.Remove(attributeLogicalName);
                }
                if (!attributeMetadata.IsValidForUpdate.Value)
                {
                    sourceAttributes.Remove(attributeLogicalName);
                }

                if (sourceAttributes.Contains(attributeLogicalName))
                {
                    tempEntity.Attributes.Add(attributeLogicalName, sourceAttributes[attributeLogicalName]);
                }
            });

            RecordToClone.Attributes.Clear();
            RecordToClone.Attributes = tempEntity.Attributes;
        }

        private bool ApplyAreAllLookupsPresentRule(CrmServiceClient TargetOrgSvc, Entity RecordToClone)
        {
            if (chkVerifyLookups.Checked)
            {
                return AreAllLookUpsPresent(TargetOrgSvc, RecordToClone);
            }

            return true;
        }

        private bool AreAllLookUpsPresent(CrmServiceClient TargetOrgSvc, Entity RecordToClone)
        {
            var missingEntityReferences = GetERefsWhichNotExistsInTargetEnvironment(TargetOrgSvc, RecordToClone);

            if (missingEntityReferences.Any())
            {
                var strBuilder = new StringBuilder();
                strBuilder.AppendLine("The following Lookups are not present in the Target Enviornment. Do you want to proceed?");

                foreach (var missingEntityReference in missingEntityReferences)
                {
                    strBuilder.AppendLine($"LogicalName {missingEntityReference.LogicalName} Id {missingEntityReference.Id} does not exist");
                }

                return MessageBox.Show(strBuilder.ToString(), "Question do you want To Procced?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            }

            return true;
        }

        private void ApplyUpsertRecords(CrmServiceClient orgService, RetrieveAllEntitiesResponse SourceMetaData, string logicalName, Entity RecordToClone)
        {
            if (chkUpsertRecords.Checked)
            {
                var primaryIdAttribute = SourceMetaData.EntityMetadata.First(entityMetaData => entityMetaData.LogicalName == logicalName).PrimaryIdAttribute;

                RecordToClone.KeyAttributes.Add(primaryIdAttribute, RecordToClone.Id);

                var upsertReq = new UpsertRequest()
                {
                    Target = RecordToClone
                };

                var resp = (UpsertResponse)orgService.Execute(upsertReq);
                var created = resp.RecordCreated;
            }
            else
            {
                orgService.Create(RecordToClone);
            }
        }

        private string GetRecordDisplayName(RetrieveAllEntitiesResponse SourceMetaData, string logicalName, Entity RecordToClone)
        {
            return RecordToClone.GetAttributeValue<string>(SourceMetaData.EntityMetadata.First(entityMetaData => entityMetaData.LogicalName == logicalName).PrimaryNameAttribute);
        }

        private RetrieveAllEntitiesResponse GetMetaData()
        {
            var metaDataRequest = new RetrieveAllEntitiesRequest
            {
                RetrieveAsIfPublished = true
            };

            var SourceMetaData = (RetrieveAllEntitiesResponse)Service.Execute(metaDataRequest);
            return SourceMetaData;
        }

        private void btnCloneRecord_Click(object sender, EventArgs e)
        {
            var recordToProcess = lstRecordsToProcess.Items;
            infoPanel = InformationPanel.GetInformationPanel(this, "Initializing...", 540, 220);
            ToggleWaitMode(true);
            using (var worker = new BackgroundWorker())
            {
                worker.DoWork += WorkerDoWorkCloneRecords;
                worker.ProgressChanged += WorkerProgressChanged;
                worker.RunWorkerCompleted += WorkerRunWorkerCompleted;
                worker.WorkerReportsProgress = true;
                worker.RunWorkerAsync(recordToProcess);
            }
        }

        private NameValueCollection GetParameterFromURL(string RecordUrl)
        {
            Uri.TryCreate(RecordUrl, UriKind.Absolute, out Uri url);
            var queryString = RecordUrl.Substring(RecordUrl.IndexOf('?')).Split('#')[0];
            var parameters = System.Web.HttpUtility.ParseQueryString(queryString); ;

            return parameters;
        }

        private object GetParameter(NameValueCollection Parameters, string Parameter)
        {
            return Parameters[Parameter];
        }

        private string GetEntityLogicalNameFromMetadataByObjectTypeCode(int typeCode)
        {
            string LogicalName;
            var metaDataRequest = new RetrieveAllEntitiesRequest
            {
                RetrieveAsIfPublished = true
            };

            var metadata = (RetrieveAllEntitiesResponse)Service.Execute(metaDataRequest);

            var targetEntity = metadata.EntityMetadata.FirstOrDefault(entity => entity.ObjectTypeCode == typeCode);

            LogicalName = targetEntity.LogicalName;
            return LogicalName;
        }

        private static void ValidDateRecordId(Guid RecordId)
        {
            if (RecordId == Guid.Empty)
            {
                throw new ArgumentException("Entity Id can't be parsed from URL");
            }
        }

        private static void ValidDateTypeCode(int typeCode)
        {
            if (typeCode == -1)
            {
                throw new ArgumentException("EntityTypeCode can't be parsed from URL");
            }
        }

        private static List<string> GetParamtersFromUrl(string query)
        {
            return query.Split('&').ToList();
        }

        private static string ReplaceQuestionMark(Uri url)
        {
            return url.Query.Replace("?", "");
        }

        private void chkVerifyLookups_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == true)
            {
                chkIgnoreAllLookups.Checked = false;
            }
        }

        private void btnAddToRecordList_Click(object sender, EventArgs e)
        {
            string RecordUrlToAdd = CleanUpRecordUrl();

            if (IsValidUrl(RecordUrlToAdd))
            {
                lstRecordsToProcess.Items.Add(RecordUrlToAdd);
                txtRecordURL.Text = string.Empty;
            }
            else
            {
                MessageBox.Show("Url is not Valid!");
            }
        }

        private static bool IsValidUrl(string fileToAdd)
        {
            return !string.IsNullOrWhiteSpace(fileToAdd);
        }

        private string CleanUpRecordUrl()
        {
            return txtRecordURL.Text.Replace("<", "").Replace(">", "");
        }

        private void lstTargetEnvironments_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnClose_Click(object sender, EventArgs e)
        {

        }

        private void lstRecordsToProcess_KeyDown(object sender, KeyEventArgs e)
        {
            if (lstRecordsToProcess.SelectedItems.Count > 0 && e.KeyCode == Keys.Delete)
            {
                lstRecordsToProcess.Items.Remove(lstRecordsToProcess.SelectedItem);
            }
        }

        protected override void ConnectionDetailsUpdated(NotifyCollectionChangedEventArgs e)
        {
            listBoxTargetEnv.DataSource = null;
            listBoxTargetEnv.DataSource = AdditionalConnectionDetails.ToList();
            listBoxTargetEnv.DisplayMember = "ConnectionName";
            listBoxTargetEnv.ValueMember = "ConnectionId";
        }
    }
}
