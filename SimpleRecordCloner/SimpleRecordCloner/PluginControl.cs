using Microsoft.Crm.Sdk.Messages;
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

namespace martintmg.MSDYN.Tools.SimpleRecordCloner
{
    public partial class PluginControl : PluginControlBase, IXrmToolBoxPluginControl, IGitHubPlugin, IStatusBarMessenger, IHelpPlugin
    {
        private ConnectionDetail detail;
        private Panel infoPanel;
        private IOrganizationService service;
        private Dictionary<string, IOrganizationService> targetServices = new Dictionary<string, IOrganizationService>();
        private KeyValuePair<string, IOrganizationService> lastTargetService;

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

        public IOrganizationService Service
        {
            get
            {
                throw new NotImplementedException();
            }
        }
        #region XrmToolbox
        public event EventHandler OnRequestConnection;
        #endregion XrmToolbox


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

        public void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName = "", object parameter = null)
        {
            this.detail = detail;
            if (actionName == "TargetOrganization")
            {
                targetServices.Add(detail.ConnectionName, newService);
                lastTargetService = new KeyValuePair<string, IOrganizationService>(detail.ConnectionName, newService);
                lstTargetEnvironments.Items.Add(new ListViewItem { Text = detail.ConnectionName });
            }
            else
            {
                service = newService;

                SetConnectionLabel(detail, "Source");

            }
        }

        #region Base tool implementation

        public PluginControl()
        {
            InitializeComponent();
        }

        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;
        public event EventHandler OnCloseTool;


        #endregion Base tool implementation

        private void btnChooseTarget_Click(object sender, EventArgs e)
        {
            if (OnRequestConnection != null)
            {
                var args = new RequestConnectionEventArgs { ActionName = "TargetOrganization", Control = this };
                OnRequestConnection(this, args);
            }
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
            if (lstTargetEnvironments.SelectedItems.Count > 0 && e.KeyCode == Keys.Delete)
            {
                foreach (ListViewItem item in lstTargetEnvironments.Items)
                {
                    if (item.Selected)
                        lstTargetEnvironments.Items.Remove(item);
                }

                ReorderTargetServices();
            }
        }

        private void ReorderTargetServices()
        {
            var oldTargetServices = targetServices;
            targetServices = new Dictionary<string, IOrganizationService>();

            foreach (ListViewItem item in lstTargetEnvironments.Items)
            {
                var serviceToAdd = oldTargetServices.First(x => x.Key == item.Text);
                targetServices.Add(serviceToAdd.Key, serviceToAdd.Value);
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
                message = "Records where cloned successfully!";
                MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void WorkerDoWorkCloneRecords(object sender, DoWorkEventArgs e)
        {
            var bw = (BackgroundWorker)sender;
            var current = 0;
            foreach (var targetService in targetServices)
            {
                var metaDataRequest = new RetrieveAllEntitiesRequest();
                metaDataRequest.RetrieveAsIfPublished = true;

                var SourceMetaData = (RetrieveAllEntitiesResponse)service.Execute(metaDataRequest);

                lastTargetService = targetService;

                var recordToProcess = (ListBox.ObjectCollection)e.Argument;

                foreach (var record in recordToProcess)
                {
                    current++;
                    string logicalName;
                    Guid id;
                    GetParameterFromURL(record.ToString(), out logicalName, out id);

                    var entity = service.Retrieve(logicalName, id, new ColumnSet(true));

                    var displayName = entity.GetAttributeValue<string>(SourceMetaData.EntityMetadata.First(entityMetaData => entityMetaData.LogicalName == logicalName).PrimaryNameAttribute);

                    bw.ReportProgress(0, string.Format("Cloning Record {4} of {5}: Displayname \"{0}\" LogicalName \"{1}\" Id \"{2}\" to Environemt \"{3}\"",
                        displayName, entity.LogicalName, entity.Id, targetService.Key,
                        current,
                        recordToProcess.Count * targetServices.Count()));
                    
                    entity.Attributes.Remove("traversedpath");

                    if (chkIgnoreAllLookups.Checked)
                    {
                        RemoveERefs(entity);
                    }

                    if (chkIgnoreOwnerAndModifiedBy.Checked)
                    {
                        RemoveOwnerAndLastmodifed(entity);
                    }

                    if (chkIgnoreStateCode.Checked)
                    {
                        RemoveStateStausCode(entity);
                    }

                    if (chkVerifyLookups.Checked)
                    {
                        if (!AreAllLookUpsPresent(entity))
                        {
                            return;
                        }
                    }

                    if (chkRemoveNotFoundLookups.Checked)
                    {
                        RemoveNotFoundLookups(entity,targetService.Value);
                    }

                    if (chkRemoveNonExistingAttributes.Checked)
                    {
                        RemoveMissingAttributeInTargetEnvironment(entity, lastTargetService.Value);
                    }

                    if (chkUpsertRecords.Checked)
                    {
                        var primaryIdAttribute = SourceMetaData.EntityMetadata.First(entityMetaData => entityMetaData.LogicalName == logicalName).PrimaryIdAttribute;

                        entity.KeyAttributes.Add(primaryIdAttribute, entity.Id);

                        var upsertReq = new UpsertRequest()
                        {
                            Target = entity
                        };

                        var resp = (UpsertResponse)lastTargetService.Value.Execute(upsertReq);
                        var created = resp.RecordCreated;
                    }
                    else
                    {
                        lastTargetService.Value.Create(entity);
                    }
                }
            }
        }

        private void RemoveNotFoundLookups(Entity Entity, IOrganizationService TargetOrgSvc)
        {
            var allERefsAttributes = Entity.Attributes.Where(attribute => attribute.Value.GetType().Name == "EntityReference").ToList();

            allERefsAttributes.ForEach(eRefAttribute =>
            {
                try
                {
                    var eRef = (EntityReference)eRefAttribute.Value;
                    var entity = TargetOrgSvc.Retrieve(eRef.LogicalName, eRef.Id, new ColumnSet(false));
                }
                catch
                {
                    Entity.Attributes.Remove(eRefAttribute.Key);
                }
            });
        }

        private void RemoveMissingAttributeInTargetEnvironment(Entity Entity, IOrganizationService TragetOrgService)
        {
            var entityMetaData = new RetrieveEntityRequest();
            entityMetaData.RetrieveAsIfPublished = true;
            entityMetaData.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes;
            entityMetaData.LogicalName = Entity.LogicalName;

            var SourceMetaData = (RetrieveEntityResponse)TragetOrgService.Execute(entityMetaData);

            var sourceAttributes = Entity.Attributes;

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

            Entity.Attributes.Clear();
            Entity.Attributes = tempEntity.Attributes;
        }

        private void RemoveStateStausCode(Entity entity)
        {
            if (entity.Contains(Helpers.EntityNames.StateCode))
                entity.Attributes.Remove(Helpers.EntityNames.StateCode);

            if (entity.Contains(Helpers.EntityNames.StatusCode))
                entity.Attributes.Remove(Helpers.EntityNames.StatusCode);
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

        private void GetParameterFromURL(string RecordUrl, out string LogicalName, out Guid RecordId)
        {
            Uri url;
            Uri.TryCreate(RecordUrl, UriKind.Absolute, out url);
            int typeCode = -1;
            var query = url.Query.Replace("?", "");
            RecordId = Guid.Empty;

            var parameters = query.Split('&').ToList();

            foreach (var parameter in parameters)
            {
                var nameAndValue = parameter.Split('=');

                if (nameAndValue[0] == "etc")
                {
                    typeCode = int.Parse(nameAndValue[1]);
                }

                if (nameAndValue[0] == "id")
                {
                    RecordId = new Guid(nameAndValue[1].Replace("%7b", "").Replace("%7d", ""));
                }
            }
            if (typeCode == -1)
            {
                throw new ArgumentException("EntityTypeCode can't be parsed from URL");
            }

            if (RecordId == Guid.Empty)
            {
                throw new ArgumentException("Entity Id can't be parsed from URL");
            }

            var metaDataRequest = new RetrieveAllEntitiesRequest();
            metaDataRequest.RetrieveAsIfPublished = true;

            var metadata = (RetrieveAllEntitiesResponse)service.Execute(metaDataRequest);

            var targetEntity = metadata.EntityMetadata.FirstOrDefault(entity => entity.ObjectTypeCode == typeCode);

            LogicalName = targetEntity.LogicalName;
        }

        private bool AreAllLookUpsPresent(Entity Entity)
        {
            var allERefsAttributes = Entity.Attributes.Where(attribute => attribute.Value.GetType().Name == "EntityReference").ToList();

            var missingEntityReferences = new List<EntityReference>();

            foreach (var eRef in allERefsAttributes)
            {
                try
                {
                    var targetEntity = lastTargetService.Value.Retrieve(((EntityReference)eRef.Value).LogicalName, ((EntityReference)eRef.Value).Id, new ColumnSet());
                }
                catch
                {
                    missingEntityReferences.Add((EntityReference)eRef.Value);
                }
            }

            if (missingEntityReferences.Any())
            {
                var strBuilder = new StringBuilder();
                strBuilder.AppendLine("The following Lookups are not present in the Target Enviornment. Do you want to proceed?");

                foreach (var missingEntityReference in missingEntityReferences)
                {
                    strBuilder.AppendLine($"LogicalName {missingEntityReference.LogicalName} Id {missingEntityReference.Id} does not exist");
                }

                return MessageBox.Show(strBuilder.ToString(), "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            }

            return true;
        }

        private static void RemoveOwnerAndLastmodifed(Entity entity)
        {
            if (entity.Contains(Helpers.EntityNames.Owner))
            {
                entity.Attributes.Remove(Helpers.EntityNames.Owner);
            }
            if (entity.Contains(Helpers.EntityNames.ModifiedBy))
            {
                entity.Attributes.Remove(Helpers.EntityNames.ModifiedBy);
            }
        }

        private static void RemoveERefs(Entity entity)
        {
            var allAttributesWithOutERefs = entity.Attributes.Where(attribute => attribute.Value.GetType().Name != "EntityReference").ToList();

            entity.Attributes = new Microsoft.Xrm.Sdk.AttributeCollection();
            entity.Attributes.AddRange(allAttributesWithOutERefs);
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
            var fileToAdd = txtRecordURL.Text.Replace("<", "").Replace(">", "");

            if (!string.IsNullOrWhiteSpace(fileToAdd))
            {
                lstRecordsToProcess.Items.Add(fileToAdd);
                txtRecordURL.Text = string.Empty;
            }
            else
            {
                MessageBox.Show("Url is not Valid!");
            }
        }

        private void lstTargetEnvironments_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void ClosingPlugin(PluginCloseInfo info)
        {
            //info.Cancel = MessageBox.Show(@"Are you sure you want to close this tab?", @"Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (OnCloseTool != null)
            {
                const string message = "Are you sure to exit?";
                if (MessageBox.Show(message, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) ==
                    DialogResult.Yes)
                    OnCloseTool(this, null);
            }
        }

        private void lstRecordsToProcess_KeyDown(object sender, KeyEventArgs e)
        {
            if (lstRecordsToProcess.SelectedItems.Count > 0 && e.KeyCode == Keys.Delete)
            {
                lstRecordsToProcess.Items.Remove(lstRecordsToProcess.SelectedItem);
            }
        }
    }
}