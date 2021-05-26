using System;
using System.Collections.Generic;
using System.Collections.Specialized;
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
using XrmToolBox.Extensibility.Interfaces;
using Label = System.Windows.Forms.Label;
using System.Drawing;

namespace MsDyn.Contrib.CloneFieldDefinitions
{
    public class CloneFieldDefinitionsControl : MultipleConnectionsPluginControlBase, IGitHubPlugin
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
        private List<EntityMetadata> _entitiesSource = new List<EntityMetadata>();
        private List<EntityMetadata> _entitiesTarget = new List<EntityMetadata>();
        private ColumnHeader columnHeader5;
        private Label label3;
        private TextBox txtPrefix;
        private Label prefixOverrideLabel;
        private CheckBox prefixOverride;
        private Button button2;
        private readonly List<EntityMetadata> _entitiesDetailedSource = new List<EntityMetadata>();
        private Button button3;
        private ColumnHeader columnHeader6;
        private List<EntityMetadata> _entitiesDetailedTarget = new List<EntityMetadata>();
        private int sortedColumn = 0;

        public string RepositoryName => "XrmToolBox.CloneFieldDefinitions";

        public string UserName => "DigitalFlow";

        public CloneFieldDefinitionsControl()
        {
            InitializeComponent();
            ConnectionUpdated += RetrieveAvailableEntities;
        }

        /// <summary>
        /// Adjust ComboBox dpopdown width dynamically based on entity names
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AdjustWidthComboBox_DropDown(object sender, System.EventArgs e)
        {
            ComboBox senderComboBox = (ComboBox)sender;
            int width = senderComboBox.DropDownWidth;
            Graphics g = senderComboBox.CreateGraphics();
            Font font = senderComboBox.Font;
            int vertScrollBarWidth =
                (senderComboBox.Items.Count > senderComboBox.MaxDropDownItems)
                ? SystemInformation.VerticalScrollBarWidth : 0;

            int newWidth;
            foreach (string s in ((ComboBox)sender).Items)
            {
                newWidth = (int)g.MeasureString(s, font).Width
                    + vertScrollBarWidth;
                if (width < newWidth)
                {
                    width = newWidth;
                }
            }
            senderComboBox.DropDownWidth = width;
            g.Dispose();
        }

        private void SetAvailableEntities()
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            _entitiesSource
                .Select(e => new { e.LogicalName, e.MetadataId })
                .OrderBy(e => e.LogicalName)
                .ToList()
                .ForEach(name =>
                {
                    comboBox1.Items.Add(name.LogicalName);
                });

            _entitiesTarget
                .Select(e => new { e.LogicalName, e.MetadataId })
                .OrderBy(e => e.LogicalName)
                .ToList()
                .ForEach(name =>
                {
                    comboBox2.Items.Add(name.LogicalName);
                });
        }

        private List<EntityMetadata> RetrieveEntities(IOrganizationService service)
        {
            var retrieveEntitiesRequest = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.Entity,
                RetrieveAsIfPublished = false
            };

            if (service == null)
            {
                throw new Exception("No Service set!");
            }

            var response = service.Execute(retrieveEntitiesRequest) as RetrieveAllEntitiesResponse;

            if (response == null)
            {
                throw new Exception("Failed to retrieve entities!");
            }

            return response.EntityMetadata.ToList();
        }

        private void RetrieveAvailableEntities(object sender, ConnectionUpdatedEventArgs eventArgs)
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving entity metadata ...",
                Work = (w, e) =>
                {
                    _entitiesSource = RetrieveEntities(Service);

                    if (AdditionalConnectionDetails.Count > 0)
                    {
                        _entitiesTarget = RetrieveEntities(GetTargetService());
                    }
                    else
                    {
                        _entitiesTarget = _entitiesSource;
                    }
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
            this.label3 = new System.Windows.Forms.Label();
            this.txtPrefix = new System.Windows.Forms.TextBox();
            this.prefixOverrideLabel = new System.Windows.Forms.Label();
            this.prefixOverride = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader6 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.button3 = new System.Windows.Forms.Button();
            this.flowLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel2.Controls.Add(this.label1);
            this.flowLayoutPanel2.Controls.Add(this.comboBox1);
            this.flowLayoutPanel2.Controls.Add(this.label2);
            this.flowLayoutPanel2.Controls.Add(this.comboBox2);
            this.flowLayoutPanel2.Controls.Add(this.button1);
            this.flowLayoutPanel2.Controls.Add(this.label3);
            this.flowLayoutPanel2.Controls.Add(this.txtPrefix);
            this.flowLayoutPanel2.Controls.Add(this.prefixOverrideLabel);
            this.flowLayoutPanel2.Controls.Add(this.prefixOverride);
            this.flowLayoutPanel2.Controls.Add(this.button2);
            this.flowLayoutPanel2.Controls.Add(this.button3);
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(1129, 39);
            this.flowLayoutPanel2.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.Margin = new Padding(8, 7, 8, 7);
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
            this.comboBox1.DropDown += new System.EventHandler(this.AdjustWidthComboBox_DropDown);            
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(206, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 13);
            this.label2.Margin = new Padding(8, 7, 8, 7);
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
            this.comboBox2.DropDown += new System.EventHandler(this.AdjustWidthComboBox_DropDown);
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Right;
            this.button1.Location = new System.Drawing.Point(406, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(136, 24);
            this.button1.TabIndex = 4;
            this.button1.Text = "Clone selected Fields";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.OnCloneButtonClick);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(548, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.Margin = new Padding(8, 7, 8, 7);
            this.label3.TabIndex = 6;
            this.label3.Text = "Field Prefix";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // txtPrefix
            // 
            this.txtPrefix.Location = new System.Drawing.Point(612, 3);
            this.txtPrefix.Name = "txtPrefix";
            this.txtPrefix.Size = new System.Drawing.Size(100, 20);
            this.txtPrefix.TabIndex = 5;
            // 
            // prefixOverrideLabel
            // 
            this.prefixOverrideLabel.AutoSize = true;
            this.prefixOverrideLabel.Location = new System.Drawing.Point(718, 0);
            this.prefixOverrideLabel.Name = "prefixOverrideLabel";
            this.prefixOverrideLabel.Size = new System.Drawing.Size(76, 13);
            this.prefixOverrideLabel.Margin = new Padding(8, 7, 8, 7);
            this.prefixOverrideLabel.TabIndex = 7;
            this.prefixOverrideLabel.Text = "Override Prefix";
            // 
            // prefixOverride
            // 
            this.prefixOverride.Location = new System.Drawing.Point(800, 3);
            this.prefixOverride.Name = "prefixOverride";
            this.prefixOverride.Size = new System.Drawing.Size(18, 24);
            this.prefixOverride.TabIndex = 8;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(824, 0);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(191, 24);
            this.button2.TabIndex = 9;
            this.button2.Text = "Target Org: Source";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // listView1
            // 
            this.listView1.Alignment = System.Windows.Forms.ListViewAlignment.SnapToGrid;
            this.listView1.AllowColumnReorder = true;
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listView1.BackColor = System.Drawing.Color.WhiteSmoke;
            this.listView1.CheckBoxes = true;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader6,
            this.columnHeader1,
            this.columnHeader3,
            this.columnHeader2,
            this.columnHeader4,
            this.columnHeader5});
            this.listView1.FullRowSelect = true;
            this.listView1.ImeMode = System.Windows.Forms.ImeMode.On;
            this.listView1.Location = new System.Drawing.Point(3, 48);
            this.listView1.Name = "listView1";
            this.listView1.OwnerDraw = true;
            this.listView1.Size = new System.Drawing.Size(1129, 443);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.ListView1_ColumnClick);
            this.listView1.DrawColumnHeader += new System.Windows.Forms.DrawListViewColumnHeaderEventHandler(this.ListView1_DrawColumnHeader);
            this.listView1.DrawItem += new System.Windows.Forms.DrawListViewItemEventHandler(this.ListView1_DrawItem);
            this.listView1.DrawSubItem += new System.Windows.Forms.DrawListViewSubItemEventHandler(this.ListView1_DrawSubItem);
            // 
            // columnHeader6
            // 
            this.columnHeader6.Width = 30;
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
            // columnHeader2
            // 
            this.columnHeader2.Text = "Type";
            this.columnHeader2.Width = 150;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "Description";
            this.columnHeader4.Width = 466;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "Custom";
            this.columnHeader5.Width = 150;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(1021, 0);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(97, 24);
            this.button3.TabIndex = 10;
            this.button3.Text = "Reset Target Org";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // CloneFieldDefinitionsControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.listView1);
            this.Name = "CloneFieldDefinitionsControl";
            this.Size = new System.Drawing.Size(1141, 500);
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        private void ListView1_DrawColumnHeader(object sender,
                                        DrawListViewColumnHeaderEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                e.DrawBackground();
                bool value = false;
                try
                {
                    value = Convert.ToBoolean(e.Header.Tag);
                }
                catch (Exception)
                {
                }
                CheckBoxRenderer.DrawCheckBox(e.Graphics,
                    new Point(e.Bounds.Left + 4, e.Bounds.Top + 4),
                    value ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal :
                    System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal);
            }
            else
            {
                e.DrawDefault = true;
            }
        }

        private void ListView1_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void ListView1_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void ListView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (e.Column == 0)
            {
                bool value = false;
                try
                {
                    value = Convert.ToBoolean(listView1.Columns[e.Column].Tag);
                }
                catch (Exception)
                {
                }
                listView1.Columns[e.Column].Tag = !value;
                foreach (ListViewItem item in listView1.Items)
                    item.Checked = !value;

                listView1.Invalidate();
            }
            else
            {
                //Check if it's the same column we sorted last
                if (sortedColumn == e.Column)
                {
                    //Alternate Sort order
                    listView1.Sorting = listView1.Sorting == SortOrder.Descending ? SortOrder.Ascending : SortOrder.Descending;
                }
                else
                {
                    //Default back to a-z sorting
                    listView1.Sorting = SortOrder.Ascending;
                }
                // Update last used column
                sortedColumn = e.Column;
                listView1.ListViewItemSorter = new ListViewItemComparer(e.Column, listView1.Sorting);
            }
        }

        private string GetDisplayLabel(Microsoft.Xrm.Sdk.Label label)
        {
            return label?.UserLocalizedLabel?.Label ?? label?.LocalizedLabels?.FirstOrDefault()?.Label;
        }

        private void OnSelectSourceEntity(object sender, EventArgs e)
        {
            comboBox2.SelectedIndex = -1;
            _entitiesDetailedSource.Clear();
            GetEntityMetadataFromServer(comboBox1.SelectedItem.ToString(), _entitiesDetailedSource, Service);
        }

        private void OnCloneButtonClick(object sender, EventArgs e)
        {
            var checkedItems = listView1.CheckedItems;

            var fieldsToClone = (from object checkedItem in checkedItems select ((ListViewItem)checkedItem).SubItems[1].Text)
                .ToList();

            var sourceEntity = (string)comboBox1.SelectedItem;
            var targetEntity = (string)comboBox2.SelectedItem;

            if (string.IsNullOrEmpty(sourceEntity) || string.IsNullOrEmpty(targetEntity))
            {
                MessageBox.Show("Source and target entity both have to be set!");
                return;
            }

            if (sourceEntity.Equals(targetEntity, StringComparison.InvariantCultureIgnoreCase) && AdditionalConnectionDetails.Count == 0)
            {
                MessageBox.Show("Source and target must not be the same!");
                return;
            }

            CloneFields(fieldsToClone, sourceEntity, targetEntity);
        }

        private EntityMetadata GetEntityMetadata(string entityName, List<EntityMetadata> metadata)
        {
            return metadata.SingleOrDefault(en => en.LogicalName == entityName);
        }

        private void GetEntityMetadataFromServer(string entityName, List<EntityMetadata> meta, IOrganizationService service)
        {
            if (meta.Any(metadata => string.Equals(metadata.LogicalName, entityName)))
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

                    if (service == null)
                    {
                        throw new Exception("No Service set!");
                    }

                    var response = service.Execute(retrieveEntityRequest) as RetrieveEntityResponse;

                    if (response == null)
                    {
                        throw new Exception("Failed to retrieve entity!");
                    }
                    meta.Add(response.EntityMetadata);
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

                    var sourceEntity = GetEntityMetadata(sourceEntityName, _entitiesDetailedSource);
                    var targetEntity = GetEntityMetadata(targetEntityName, _entitiesDetailedTarget);

                    var errors = new Dictionary<string, Exception>();

                    for (var i = 0; i < fields.Count; i++)
                    {
                        var fieldName = fields[i];
                        AttributeMetadata targetAttribute = null;
                        var attribute = sourceEntity.Attributes.Single(attr => attr.LogicalName.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase));
                        if (!attribute.IsCustomAttribute.Value)
                        {
                            targetAttribute = targetEntity.Attributes.SingleOrDefault(attr => attr.LogicalName.Equals(String.Format("{0}_{1}", txtPrefix.Text.Replace("_", ""), fieldName).ToLower(), StringComparison.InvariantCultureIgnoreCase));
                        }
                        else
                        {
                            targetAttribute = targetEntity.Attributes.SingleOrDefault(attr => attr.LogicalName.Equals(fieldName, StringComparison.InvariantCultureIgnoreCase));
                        }

                        if (targetAttribute != null)
                        {
                            continue;
                        }

                        // Don't regenerate ID if transfering to different organization
                        if (AdditionalConnectionDetails == null || AdditionalConnectionDetails.Count == 0)
                        {
                            attribute.MetadataId = Guid.NewGuid();
                        }

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

                        var textPrefix = txtPrefix.Text.Replace("_", "");
                        if (!attribute.IsCustomAttribute.Value)
                        {
                            attribute.LogicalName = String.Format("{0}_{1}", textPrefix, attribute.LogicalName).ToLower();
                            attribute.SchemaName = String.Format("{0}_{1}", textPrefix, attribute.SchemaName).ToLower();
                        }

                        if (this.prefixOverride.Checked)
                        {
                            attribute.LogicalName = String.Format("{0}_{1}", textPrefix, attribute.LogicalName.Substring(attribute.LogicalName.IndexOf('_') + 1)).ToLower();
                            attribute.SchemaName = String.Format("{0}_{1}", textPrefix, attribute.SchemaName.Substring(attribute.SchemaName.IndexOf('_') + 1)).ToLower();
                        }

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
                            if (AdditionalConnectionDetails.Count > 0)
                            {
                                AdditionalConnectionDetails[0].GetCrmServiceClient().Execute(request);
                            }
                            else
                            {
                                Service.Execute(request);
                            }
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
            {
                if (comboBox2.SelectedItem != null)
                {
                    GetEntityMetadataFromServer(comboBox2.SelectedItem.ToString(), _entitiesDetailedTarget, GetTargetService());
                }
            }
        }

        private void PopulateAttributeList()
        {
            listView1.Items.Clear();

            var sourceEntity = _entitiesDetailedSource.SingleOrDefault(en => en.LogicalName == (string)comboBox1.SelectedItem);
            var targetEntity = _entitiesDetailedTarget.SingleOrDefault(en => en.LogicalName == (string)comboBox2.SelectedItem);

            if (sourceEntity == null || targetEntity == null)
            {
                return;
            }

            var attributes = sourceEntity.Attributes
                // Get only custom attributes and leave out base fields for currencies
                .Where(attr =>
                    attr.IsCustomAttribute.HasValue
                    //&& attr.IsCustomAttribute.Value 
                    && attr.IsValidForCreate.HasValue &&
                    attr.IsValidForCreate.Value)
                .Where(attr => targetEntity == null || targetEntity.Attributes.All(a => !a.LogicalName.Equals(attr.LogicalName)))
                .ToList();

            foreach (var attributeMetadata in attributes)
            {
                var item = new ListViewItem();
                item.SubItems.Add(attributeMetadata.LogicalName);
                item.SubItems.Add(GetDisplayLabel(attributeMetadata.DisplayName));
                item.SubItems.Add(attributeMetadata.AttributeType.ToString());
                item.SubItems.Add(GetDisplayLabel(attributeMetadata.Description));
                item.SubItems.Add(attributeMetadata.IsCustomAttribute.Value.ToString());

                listView1.Items.Add(item);
            }
        }

        private IOrganizationService GetTargetService()
        {
            if (AdditionalConnectionDetails.Count > 0)
            {
                return AdditionalConnectionDetails[0].GetCrmServiceClient();
            }
            else
            {
                if (Service == null)
                {
                    throw new Exception("No Service set!");
                }

                return Service;
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        protected override void ConnectionDetailsUpdated(NotifyCollectionChangedEventArgs e)
        {
            _entitiesDetailedTarget = new List<EntityMetadata>();
            RetrieveAvailableEntities(null, null);
            button2.Text = $"Target Org: {(AdditionalConnectionDetails != null && AdditionalConnectionDetails.Count > 0 ? AdditionalConnectionDetails[0].Organization : "Source")}";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (AdditionalConnectionDetails != null && AdditionalConnectionDetails.Count > 0)
            {
                MessageBox.Show("Only one additional connection is allowed");
                return;
            }

            AddAdditionalOrganization();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(AdditionalConnectionDetails == null || AdditionalConnectionDetails.Count == 0)
            {
                return;
            }

            RemoveAdditionalOrganization(AdditionalConnectionDetails[0]);
            button2.Text = "Target Org: Source";
            RetrieveAvailableEntities(null, null);
        }
    }
}
