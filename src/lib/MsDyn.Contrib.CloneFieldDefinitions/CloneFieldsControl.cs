using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using XrmToolBox.Extensibility;
using Label = System.Windows.Forms.Label;

namespace MsDyn.Contrib.CloneFieldDefinitions
{
    public class CloneFieldDefinitionsControl : PluginControlBase
    {
        private FlowLayoutPanel flowLayoutPanel2;
        private Label label1;
        private ComboBox comboBox1;
        private Label label2;
        private ComboBox comboBox2;
        private ListView listView1;
        private ColumnHeader columnHeader1;
        private ColumnHeader columnHeader2;
        private ColumnHeader columnHeader3;
        private ColumnHeader columnHeader4;
        private Button button1;
        private List<EntityMetadata> _entities;
        private readonly List<EntityMetadata> _entitiesDetailed;


        public CloneFieldDefinitionsControl()
        {
            InitializeComponent();
            _entitiesDetailed = new List<EntityMetadata>();
            ConnectionUpdated += RetrieveAvailableEntities;
        }

        private void SetAvailableEntities()
        {
            _entities
                .Select(e => e.LogicalName)
                .OrderBy(e => e)
                .ToList()
                .ForEach(name =>
                {
                    comboBox1.Items.Add(name);
                    comboBox2.Items.Add(name);
                });
        }

        private void RetrieveAvailableEntities(object sender, ConnectionUpdatedEventArgs eventArgs)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving entity metadata ...",
                Work = (w, e) =>
                {
                    var retrieveEntitiesRequest = new RetrieveAllEntitiesRequest
                    {
                        EntityFilters = EntityFilters.Entity,
                        RetrieveAsIfPublished = false
                    };

                    if (Service == null)
                    {
                        throw new Exception("No Service set!");
                    }

                    var response = Service.Execute(retrieveEntitiesRequest) as RetrieveAllEntitiesResponse;

                    if (response == null)
                    {
                        throw new Exception("Failed to retrieve entities!");
                    }

                    _entities = response.EntityMetadata.ToList();
                },
                PostWorkCallBack = e =>
                {
                    SetAvailableEntities();
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private void InitializeComponent()
        {
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.flowLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Controls.Add(this.label1);
            this.flowLayoutPanel2.Controls.Add(this.comboBox1);
            this.flowLayoutPanel2.Controls.Add(this.label2);
            this.flowLayoutPanel2.Controls.Add(this.comboBox2);
            this.flowLayoutPanel2.Controls.Add(this.button1);
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(920, 39);
            this.flowLayoutPanel2.TabIndex = 0;
            this.flowLayoutPanel2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source Entity";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(79, 3);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 1;
            this.comboBox1.SelectedValueChanged += new System.EventHandler(this.OnSelectSourceEntity);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(206, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Target Entity";
            // 
            // comboBox2
            // 
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(279, 3);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(121, 21);
            this.comboBox2.TabIndex = 3;
            this.comboBox2.SelectedValueChanged += new System.EventHandler(this.OnSelectTargetEntity);
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Right;
            this.button1.Location = new System.Drawing.Point(406, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(136, 21);
            this.button1.TabIndex = 4;
            this.button1.Text = "Clone selected Fields";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.OnCloneButtonClick);
            // 
            // listView1
            // 
            this.listView1.Alignment = System.Windows.Forms.ListViewAlignment.SnapToGrid;
            this.listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent | ColumnHeaderAutoResizeStyle.HeaderSize);
            this.listView1.HeaderStyle = ColumnHeaderStyle.Clickable;
            this.listView1.AllowColumnReorder = true;
            this.listView1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.listView1.CheckBoxes = true;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader3,
            this.columnHeader2,
            this.columnHeader4});
            this.listView1.FullRowSelect = true;
            this.listView1.Location = new System.Drawing.Point(3, 48);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(920, 443);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                    | System.Windows.Forms.AnchorStyles.Left)
                    | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Logical Name";
            this.columnHeader1.Width = 150;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Display Name";
            this.columnHeader3.Width = 150;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Description";
            this.columnHeader4.Width = -2;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Type";
            this.columnHeader2.Width = 150;
            // 
            // CloneFieldDefinitionsControl
            // 
            this.AutoSize = true;
            this.Controls.AddRange(new Control[] { this.flowLayoutPanel2, this.listView1 });
            this.Name = "CloneFieldDefinitionsControl";
            this.Size = new System.Drawing.Size(932, 500);
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private string GetDisplayLabel(Microsoft.Xrm.Sdk.Label label)
        {
            return label?.UserLocalizedLabel?.Label ?? label?.LocalizedLabels?.FirstOrDefault()?.Label;
        }

        private void OnSelectSourceEntity(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = -1;
            _entitiesDetailed.Clear();
            GetEntityMetadataFromServer(comboBox1.SelectedItem.ToString());
        }

        private void OnCloneButtonClick(object sender, EventArgs e)
        {
            var checkedItems = listView1.CheckedItems;

            var fieldsToClone = (from object checkedItem in checkedItems select ((ListViewItem)checkedItem).Text)
                .ToList();

            var sourceEntity = (string)comboBox1.SelectedItem;
            var targetEntity = (string)comboBox2.SelectedItem;

            if (string.IsNullOrEmpty(sourceEntity) || string.IsNullOrEmpty(targetEntity))
            {
                MessageBox.Show("Source and target entity both have to be set!");
                return;
            }

            if (sourceEntity.Equals(targetEntity, StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show("Source and target must not be the same!");
                return;
            }

            CloneFields(fieldsToClone, sourceEntity, targetEntity);
        }

        private EntityMetadata GetEntityMetadata(string entityName)
        {
            return _entitiesDetailed.SingleOrDefault(en => en.LogicalName == entityName);
        }

        private void GetEntityMetadataFromServer(string entityName)
        {
            if (_entitiesDetailed.Any(metadata => string.Equals(metadata.LogicalName, entityName)))
            {
                PopulateAttributeList();
                return;
            }
            WorkAsync(new WorkAsyncInfo
            {
                Message = $"Retrieving {entityName} metadata ...",
                Work = (w, e) =>
                {
                    var retrieveEntityRequest = new RetrieveEntityRequest
                    {
                        LogicalName = entityName,
                        EntityFilters = EntityFilters.All,
                        RetrieveAsIfPublished = false
                    };

                    if (Service == null)
                    {
                        throw new Exception("No Service set!");
                    }

                    var response = Service.Execute(retrieveEntityRequest) as RetrieveEntityResponse;

                    if (response == null)
                    {
                        throw new Exception("Failed to retrieve entity!");
                    }
                    _entitiesDetailed.Add(response.EntityMetadata);
                },
                PostWorkCallBack = e =>
                {
                    PopulateAttributeList();
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private void CloneFields(List<string> fields, string sourceEntityName, string targetEntityName)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Cloning fields ...",
                Work = (w, e) =>
                {
                    w.WorkerReportsProgress = true;

                    var sourceEntity = GetEntityMetadata(sourceEntityName);
                    var targetEntity = GetEntityMetadata(targetEntityName);

                    var errors = new Dictionary<string, Exception>();

                    for (var i = 0; i < fields.Count; i++)
                    {
                        var fieldName = fields[i];

                        var attribute = sourceEntity.Attributes.Single(attr => attr.LogicalName.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase));
                        var targetAttribute = targetEntity.Attributes.SingleOrDefault(attr => attr.LogicalName.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase));

                        if (targetAttribute != null)
                        {
                            continue;
                        }

                        attribute.MetadataId = Guid.NewGuid();

                        CloneBooleanAttribute(sourceEntityName, targetEntityName, attribute);
                        CloneOptionSetAttribute(sourceEntityName, targetEntityName, attribute);
                        CloneMultipleLinesOfTextAttribute(sourceEntityName, targetEntityName, attribute);

                        OrganizationRequest request;

                        var isEntityReference = new List<AttributeTypeCode>
                        {
                            AttributeTypeCode.Customer,
                            AttributeTypeCode.Lookup,
                            AttributeTypeCode.Owner
                        }.Contains(attribute.AttributeType.Value);

                        if (!isEntityReference)
                        {
                            request = new CreateAttributeRequest
                            {
                                Attribute = attribute,
                                EntityName = targetEntityName
                            };
                        }
                        else
                        {
                            request = CloneLookupAttribute(sourceEntity, targetEntity, attribute)
                                ?? CloneCustomerAttribute(sourceEntity, targetEntity, attribute);
                        }

                        try
                        {
                            Service.Execute(request);
                        }
                        catch (Exception ex)
                        {
                            errors.Add(fieldName, ex);
                        }

                        w.ReportProgress((int)((i + 1) / (double)fields.Count * 100));
                    }

                    e.Result = errors;
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage($"Progress: {e.ProgressPercentage} %");
                },
                PostWorkCallBack = e =>
                {
                    var faults = (Dictionary<string, Exception>)e.Result;

                    if (faults.Count == 0)
                    {
                        MessageBox.Show("Success!");
                    }
                    else
                    {
                        MessageBox.Show($"Error: {string.Join(", ", faults.Select(r => $"{r.Key}: {r.Value.Message}"))}");
                    }
                },
                AsyncArgument = null,
                IsCancelable = true,
                MessageWidth = 340,
                MessageHeight = 150
            });
        }

        private void CloneMultipleLinesOfTextAttribute(string sourceEntityName, string targetEntityName, AttributeMetadata attribute)
        {
            var memoAttribute = attribute as MemoAttributeMetadata;

            if (memoAttribute == null)
            {
                return;
            }

            memoAttribute.Format = null;
        }

        private OrganizationRequest CloneCustomerAttribute(EntityMetadata sourceEntity, EntityMetadata targetEntity, AttributeMetadata attribute)
        {
            return null;
        }

        private OrganizationRequest CloneLookupAttribute(EntityMetadata sourceEntity, EntityMetadata targetEntity, AttributeMetadata attribute)
        {
            var lookupAttribute = attribute as LookupAttributeMetadata;

            if (lookupAttribute == null)
            {
                return null;
            }

            var relationShip =
                sourceEntity.ManyToOneRelationships.SingleOrDefault(
                    rel => rel.ReferencingAttribute.Equals(lookupAttribute.LogicalName, StringComparison.InvariantCultureIgnoreCase));

            if (relationShip == null)
            {
                return null;
            }

            relationShip.ReferencingEntity = targetEntity.LogicalName;
            relationShip.ReferencingAttribute = string.Empty;

            relationShip.ReferencedEntityNavigationPropertyName =
                relationShip.ReferencedEntityNavigationPropertyName.ReplaceEntityName(sourceEntity.LogicalName, targetEntity.LogicalName);

            relationShip.SchemaName = relationShip.SchemaName.ReplaceEntityName(sourceEntity.LogicalName, targetEntity.LogicalName);

            var lookup = new LookupAttributeMetadata
            {
                Description = lookupAttribute.Description,
                DisplayName = lookupAttribute.DisplayName,
                LogicalName = lookupAttribute.LogicalName,
                SchemaName = lookupAttribute.SchemaName,
                RequiredLevel = lookupAttribute.RequiredLevel
            };

            var request = new CreateOneToManyRequest
            {
                Lookup = lookup,
                OneToManyRelationship = relationShip
            };

            return request;
        }

        private void CloneOptionSetAttribute(string sourceEntityName, string targetEntityName, AttributeMetadata attribute)
        {
            var optionSetAttribute = attribute as PicklistAttributeMetadata;
            if (optionSetAttribute?.OptionSet?.IsGlobal != null &&
                !optionSetAttribute.OptionSet.IsGlobal.Value)
            {
                optionSetAttribute.OptionSet.MetadataId = Guid.NewGuid();
                optionSetAttribute.OptionSet.Name = optionSetAttribute.OptionSet.Name.ReplaceEntityName(sourceEntityName, targetEntityName);
            }
            else if (optionSetAttribute?.OptionSet?.IsGlobal != null &&
                optionSetAttribute.OptionSet.IsGlobal.Value)
            {
                optionSetAttribute.OptionSet.MetadataId = Guid.NewGuid();
                optionSetAttribute.OptionSet.IsGlobal = true;
                optionSetAttribute.OptionSet.Name = optionSetAttribute.OptionSet.Name.Replace(sourceEntityName,
                    targetEntityName);
                optionSetAttribute.OptionSet.Options.Clear();

            }
        }

        private void CloneBooleanAttribute(string sourceEntityName, string targetEntityName, AttributeMetadata attribute)
        {
            var booleanAttribute = attribute as BooleanAttributeMetadata;

            if (booleanAttribute?.OptionSet?.IsGlobal != null && !booleanAttribute.OptionSet.IsGlobal.Value)
            {
                booleanAttribute.OptionSet.MetadataId = Guid.NewGuid();
                booleanAttribute.OptionSet.Name = booleanAttribute.OptionSet.Name.ReplaceEntityName(sourceEntityName, targetEntityName);
            }
        }

        private void OnSelectTargetEntity(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex < 0)
                MessageBox.Show("Please Choose Source Entity.");
            else
                if (comboBox2.SelectedItem != null)
                GetEntityMetadataFromServer(comboBox2.SelectedItem.ToString());
        }

        private void PopulateAttributeList()
        {
            listView1.Items.Clear();

            var sourceEntity = _entitiesDetailed.SingleOrDefault(en => en.LogicalName == (string)comboBox1.SelectedItem);
            var targetEntity = _entitiesDetailed.SingleOrDefault(en => en.LogicalName == (string)comboBox2.SelectedItem);

            if (sourceEntity == null)
            {
                return;
            }

            var attributes = sourceEntity.Attributes
                // Get only custom attributes and leave out base fields for currencies
                .Where(attr =>
                    attr.IsCustomAttribute.HasValue && attr.IsCustomAttribute.Value && attr.IsValidForCreate.HasValue &&
                    attr.IsValidForCreate.Value)
                .Where(attr => targetEntity == null || targetEntity.Attributes.All(a => !a.LogicalName.Equals(attr.LogicalName)))
                .ToList();

            foreach (var attributeMetadata in attributes)
            {
                var item = new ListViewItem(attributeMetadata.LogicalName);

                item.SubItems.Add(GetDisplayLabel(attributeMetadata.DisplayName));
                item.SubItems.Add(attributeMetadata.AttributeType.ToString());
                item.SubItems.Add(GetDisplayLabel(attributeMetadata.Description));

                listView1.Items.Add(item);
            }
        }
    }
}
