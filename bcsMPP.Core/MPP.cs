using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Xml;
using Aras.IOM;
using bwInnovatorCore;
using System.Web;
using System.IO;

namespace bcsMPP.Core
{
    public class MPP
    {
        #region "                   宣告區"

        protected Innovator Cinn;
        protected CInnovator.bwGeneric CbwGeneric = new CInnovator.bwGeneric(); //Modify by kenny 2019/04/11
        protected CGeneric.Common CoCommon;//Modify by kenny 2019/04/11
        private string CstrErrMessage = "";

        protected Innovator innovator { get; private set; }
        protected string LangCode { get; private set; }
        #endregion

        #region "                   進入點"

        public MPP()
        {
            Cinn = new Innovator(null);
        }


        public MPP(Innovator getInnovator)
        {
            //System.Diagnostics.Debugger.Break();

            innovator = getInnovator;
            Cinn = getInnovator;
            CbwGeneric.bwIOMInnovator = Cinn;
            CoCommon = new CGeneric.Common();

            //Modify by kenny 2016/04/01 ------
            string LanCode = Cinn.getI18NSessionContext().GetLanguageCode();
            if (LanCode == null) LanCode = "";
            LangCode = LanCode;
            LanCode = LanCode.ToLower();

            if ((LanCode.IndexOf("zt") > -1) || (LanCode.IndexOf("tw") > -1))
            {
                LanCode = "zh-tw";
            }
            else if ((LanCode.IndexOf("zc") > -1) || (LanCode.IndexOf("cn") > -1))
            {
                LanCode = "zh-cn";
            }
            else if ((LanCode.IndexOf("kr") > -1) || (LanCode.IndexOf("ko") > -1))
            {
                LanCode = "ko-kr";
            }
            else
            {
                LanCode = "en";
            }
            CoCommon.SetLanguage = LanCode;
            CbwGeneric.SetLanguage = LanCode;
            //----------------------------------
        }

        #endregion

        #region "                   屬性區"

        protected virtual string CstrLicenseCode
        {
            get { return "AO-09011"; }
        }

        //Add Property by kenny 2016/03/02
        public string ErrMessage
        {
            get { return CstrErrMessage; }
        }


        #endregion

        #region "                   方法區"

        /// <summary>
        /// 读取MPP ProcessPlan数据
        /// </summary>
        /// <param name="mpp">MPP对象</param>
        /// <returns>返回ProcessPlan数据</returns>
        public Item getProcessPlanStructure(Item mpp)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            Innovator innovatorInstance = this.innovator;
            Item query = innovatorInstance.newItem("mpp_PartConfig", "get");
            query.setID("38F55BAEEA1B49BCA95B472BF4D9C133");
            query.setAttribute("select", "default_name_p_name,default_item_number_p_name,name_p_name,item_number_p_name");
            Item partConfig = query.apply();
            if (partConfig.isError())
            {
                return partConfig;
            }
            string nameProperty = partConfig.getProperty("name_p_name");
            if (string.IsNullOrEmpty(nameProperty))
            {
                nameProperty = partConfig.getProperty("default_name_p_name");
            }
            string itemNumberProperty = partConfig.getProperty("item_number_p_name");
            if (string.IsNullOrEmpty(itemNumberProperty))
            {
                itemNumberProperty = partConfig.getProperty("default_item_number_p_name");
            }

            //Modify By tengz 2020/1/19
            //调整MPP ProcessPlan Tree Root显示为MPP名称+Part KeyedName
            string customPartProperties = nameProperty + "," + itemNumberProperty + ",keyed_name"; ;

            //Modify by tengz 2020/1/20
            //动态读取QuickEditForm属性
            List<string> propertys = getQuickEditFormPropertys(innovatorInstance);

            string langCode = mpp.getProperty("lang_code");
            Item foundProcessPlans = innovatorInstance.newItem();

            Item processPlanItem = getProcessPlanStructure(innovatorInstance, mpp.getProperty("with_details_ids", ""), true, langCode, customPartProperties, mpp.getProperty("bcs_location", ""), propertys);
            if (processPlanItem != null && !processPlanItem.isError())
            {
                foundProcessPlans.appendItem(processPlanItem);
            }

            processPlanItem = getProcessPlanStructure(innovatorInstance, mpp.getProperty("plan_ids", ""), false, langCode, customPartProperties, mpp.getProperty("bcs_location", ""), propertys);
            if (processPlanItem != null && !processPlanItem.isError())
            {
                foundProcessPlans.appendItem(processPlanItem);
            }

            if (foundProcessPlans.getItemCount() == 1)
            {
                foundProcessPlans = innovatorInstance.newResult(String.Empty);
            }
            else
            {
                // cleaning up item, that created during variable declaration
                foundProcessPlans.removeItem(foundProcessPlans.getItemByIndex(0));

                // cut off all item nodes with related items, which hove no 'get' access
                XmlNodeList restrictedItemNodes = foundProcessPlans.dom.SelectNodes("//Item[related_id[@is_null='0' or @discover_only='1']]");
                foreach (XmlNode itemNode in restrictedItemNodes)
                {
                    if (itemNode.ParentNode != null)
                    {
                        itemNode.ParentNode.RemoveChild(itemNode);
                    }
                }

                //Modify by BCS Tengz 2021/6/23 MPP与PQD联动 读取PQD数据
                if (CheckIsUsedPQD())
                {
                    foundProcessPlans = getRelationPQDData(foundProcessPlans);
                }
                //End Modify
            }

            return foundProcessPlans;
        }

        static Item getProcessPlanStructure(Innovator innovatorInstance, string processPlanIds, bool withDetails, string langCode, string customPartProperties, string location, List<string> propertys)
        {
            if (!String.IsNullOrEmpty(processPlanIds))
            {
                //Modify by tengz 2019/6/4
                //添加编辑器内容存储属性(bcs_details)
                //Modify by tengz 2019/6/4
                //添加Location处理
                //Modify By tengz 2020/2/4
                //调整MPP ProcessPlan Tree 显示对象KeyedName
                string bcs_location = !string.IsNullOrEmpty(location) ? "<OR><bcs_location>" + location + "</bcs_location><bcs_location condition='is null'></bcs_location></OR>" : string.Empty;

                String requestAml = @"<Item type='mpp_ProcessPlan' action='get' select='bcs_head,bcs_foot,name,item_number,location,description,locked_by_id' idlist='" + processPlanIds + (!string.IsNullOrEmpty(langCode) ? "' language='" + langCode : string.Empty) + @"'>
					<Relationships>
						<Item type='mpp_ProcessPlanProducedPart' select='related_id(" + customPartProperties + @",classification)'/>
						<Item type='mpp_ProcessPlanLocation' select='related_id(name,item_number)'/>
						<Item type='mpp_Operation' select='bcs_template,bcs_type,bcs_hours,bcs_location,keyed_name,name,sort_order" + propertys[0] + (withDetails ? ",wi_details,bcs_details" : String.Empty) + @"'" + (!string.IsNullOrEmpty(langCode) ? " language='" + langCode + "'" : string.Empty) + @" >"
                                + bcs_location + @"
							<Relationships>" +
                                    (withDetails ?
                                    @"<Item type='mpp_Step' " + (!string.IsNullOrEmpty(langCode) ? " language='" + langCode + "'" : string.Empty) + @">"
                                        + bcs_location + @"
									<Relationships>
										<Item type='mpp_StepImageReference' select='reference_id,related_id(src)'/>
									</Relationships>
								</Item>"
                                    : "<Item type='mpp_Step' select='bcs_location,keyed_name,name,sort_order" + propertys[1] + "'>" + bcs_location + "</Item>") +
                                    @"<Item type='mpp_OperationConsumedPart' select='bcs_location,quantity,related_id(" + customPartProperties + @",classification,config_id)'>'" + bcs_location + @"</Item>
								<Item type='mpp_OperationResource' select='bcs_location,related_id(keyed_name,name,item_number)'>" + bcs_location + @"</Item>
								<Item type='mpp_OperationSkill' select='bcs_location,related_id(keyed_name,name,item_number)'>" + bcs_location + @"</Item>
								<Item type='mpp_OperationCAD' select='bcs_location,related_id(keyed_name,name,item_number,major_rev)'>" + bcs_location + @"</Item>
								<Item type='mpp_OperationTest' select='bcs_location,related_id(keyed_name,name,item_number)'>" + bcs_location + @"</Item>
								<Item type='mpp_OperationDocument' select='bcs_location,related_id(keyed_name,name,item_number,major_rev)'>" + bcs_location + "</Item>" +
                                    (withDetails ? "<Item type='mpp_OperationImageReference' select='reference_id,related_id(src)'/>" : "") +
                                @"</Relationships>
						</Item>
					</Relationships>
				</Item>";

                Item requestItem = innovatorInstance.newItem();
                requestItem.loadAML(requestAml);

                return requestItem.apply();
            }

            return null;
        }

        /// <summary>
        /// 读取关联PQD数据
        /// </summary>
        /// <param name="foundProcessPlans"></param>
        /// <returns></returns>
        private Item getRelationPQDData(Item foundProcessPlans)
        {
            int _count = foundProcessPlans.getItemCount();
            string _aml = "     <AML>" +
                          "       <Item type='Process Quality Document' action='get' id='{0}' select='id' >" +
                          "         <Relationships>" +
                          "            <Item type='PQD Operation' action='get' select='reference_id, parent_reference_id, sort_order, bound_item_id, bound_item_config_id, tracking_mode, resolution_mode' />" +
                          "            <Item type='PQD EMT' action='get' select='reference_id, parent_reference_id, sort_order, bound_item_id, bound_item_config_id, tracking_mode, resolution_mode' />" +
                          "            <Item type='PQD Tool' action='get' select='reference_id, parent_reference_id, sort_order, bound_item_id, bound_item_config_id, tracking_mode, resolution_mode' />" +
                          "         </Relationships>" +
                          "       </Item>" +
                          "     </AML>";
            Item requestItem = innovator.newItem("Method", "cmf_GetItemsAndOptimizeResult");
            for (int i = 0; i < _count; i++)
            {
                Item mppItem = foundProcessPlans.getItemByIndex(i);
                Item pqdItem = innovator.newItem("Process Quality Document", "get");
                pqdItem.setProperty("process_plan_id", mppItem.getID());
                pqdItem.setAttribute("select", "id");
                pqdItem = pqdItem.apply();
                if (pqdItem.getItemCount() != 1)
                {
                    continue;
                }
                requestItem.setProperty("amlGetQuery", string.Format(_aml, pqdItem.getID()));
                requestItem = requestItem.apply();
                if (requestItem.isError() || requestItem.getItemCount() != 1)
                {
                    continue;
                }

                XmlNode cmfPQDdata = requestItem.node.SelectSingleNode("Rels_someSaltGFJHpiwy3");
                if (cmfPQDdata == null)
                {
                    continue;
                }

                Item operationItems = mppItem.getRelationships("mpp_Operation");
                int _opCount = operationItems.getItemCount();
                for (int j = 0; j < _opCount; j++)
                {
                    Item operationItem = operationItems.getItemByIndex(j);
                    XmlNode cmfOPItem = cmfPQDdata.SelectSingleNode($"G[@e='PQD Operation']/I[@f='{operationItem.getID()}']");
                    if (cmfOPItem == null)
                    {
                        continue;
                    }

                    //处理检验设备
                    List<Item> toolItems = new List<Item>();
                    XmlNodeList cmfToolItems = cmfPQDdata.SelectNodes($"G[@e='PQD Tool']/I[@c='{cmfOPItem.Attributes["a"].Value}']");
                    foreach (XmlNode cmfToolItem in cmfToolItems)
                    {
                        Item toolItem = innovator.getItemById("mpp_Resource", cmfToolItem.Attributes["f"].Value);
                        if (toolItem.isError())
                        {
                            continue;
                        }
                        Item testToolItem = innovator.newItem("mpp_OperationTest");
                        testToolItem.setRelatedItem(toolItem);
                        testToolItem.setID(cmfToolItem.Attributes["d"].Value);
                        toolItems.Add(testToolItem);
                    }

                    //处理检验项目
                    XmlNodeList cmfTestItems = cmfPQDdata.SelectNodes($"G[@e='PQD EMT']/I[@c='{cmfOPItem.Attributes["a"].Value}']");
                    foreach (XmlNode cmfTestItem in cmfTestItems)
                    {
                        Item testItem = innovator.getItemById("mpp_Test", cmfTestItem.Attributes["f"].Value);
                        if (testItem.isError())
                        {
                            continue;
                        }

                        Item testRelItem = operationItem.createRelationship("mpp_OperationTest", "skip");
                        testRelItem.setRelatedItem(testItem);
                        //testRelItem.setNewID();
                        testRelItem.setID(cmfTestItem.Attributes["d"].Value);

                        foreach (Item toolItem in toolItems)
                        {
                            //因MPP程式逻辑原因此处关系要加在mpp_OperationTest下而不是mpp_Test下
                            testRelItem.addRelationship(toolItem);
                        }
                    }
                }
            }

            return foundProcessPlans;
        }

        private Item getRelationPQDData(string mppId, Item operationItems,ref Dictionary<string,string> itemTypeIcons)
        {
            string _aml = "     <AML>" +
                          "       <Item type='Process Quality Document' action='get' id='{0}' select='id' >" +
                          "         <Relationships>" +
                          "            <Item type='PQD Operation' action='get' select='reference_id, parent_reference_id, sort_order, bound_item_id, bound_item_config_id, tracking_mode, resolution_mode' />" +
                          "            <Item type='PQD EMT' action='get' select='reference_id, parent_reference_id, sort_order, bound_item_id, bound_item_config_id, tracking_mode, resolution_mode' />" +
                          "            <Item type='PQD Tool' action='get' select='reference_id, parent_reference_id, sort_order, bound_item_id, bound_item_config_id, tracking_mode, resolution_mode' />" +
                          "         </Relationships>" +
                          "       </Item>" +
                          "     </AML>";
            Item requestItem = innovator.newItem("Method", "cmf_GetItemsAndOptimizeResult");
            Item pqdItem = innovator.newItem("Process Quality Document", "get");
            pqdItem.setProperty("process_plan_id", mppId);
            pqdItem.setAttribute("select", "id");
            pqdItem = pqdItem.apply();
            if (pqdItem.getItemCount() != 1)
            {
                return operationItems;
            }
            requestItem.setProperty("amlGetQuery", string.Format(_aml, pqdItem.getID()));
            requestItem = requestItem.apply();
            if (requestItem.isError() || requestItem.getItemCount() != 1)
            {
                return operationItems;
            }

            XmlNode cmfPQDdata = requestItem.node.SelectSingleNode("Rels_someSaltGFJHpiwy3");
            if (cmfPQDdata == null)
            {
                return operationItems;
            }

            int _opCount = operationItems.getItemCount();
            for (int j = 0; j < _opCount; j++)
            {
                Item operationItem = operationItems.getItemByIndex(j);
                XmlNode cmfOPItem = cmfPQDdata.SelectSingleNode($"G[@e='PQD Operation']/I[@f='{operationItem.getID()}']");
                if (cmfOPItem == null)
                {
                    continue;
                }

                //处理检验设备
                List<Item> toolItems = new List<Item>();
                XmlNodeList cmfToolItems = cmfPQDdata.SelectNodes($"G[@e='PQD Tool']/I[@c='{cmfOPItem.Attributes["a"].Value}']");
                foreach (XmlNode cmfToolItem in cmfToolItems)
                {
                    Item toolItem = innovator.getItemById("mpp_Resource", cmfToolItem.Attributes["f"].Value);
                    if (toolItem.isError())
                    {
                        continue;
                    }
                    Item ToolRelItem = innovator.newItem("mpp_TestTool");
                    ToolRelItem.setRelatedItem(toolItem);
                    //ToolRelItem.setNewID();
                    ToolRelItem.setID(cmfToolItem.Attributes["d"].Value);
                    toolItems.Add(ToolRelItem);

                    string itemTypeName = toolItem.getType();
                    if (!itemTypeIcons.ContainsKey(itemTypeName))
                    {
                        Item itemType = innovator.newItem("ItemType", "get");
                        itemType.setProperty("name", itemTypeName);
                        itemType.setAttribute("select", "");
                        itemType = itemType.apply();
                        itemTypeIcons.Add(itemTypeName, itemType.getProperty("open_icon", "../images/ItemType.svg"));
                    }
                }

                //处理检验项目
                XmlNodeList cmfTestItems = cmfPQDdata.SelectNodes($"G[@e='PQD EMT']/I[@c='{cmfOPItem.Attributes["a"].Value}']");
                foreach (XmlNode cmfTestItem in cmfTestItems)
                {
                    Item testItem = innovator.getItemById("mpp_Test", cmfTestItem.Attributes["f"].Value);
                    if (testItem.isError())
                    {
                        continue;
                    }

                    foreach (Item toolItem in toolItems)
                    {
                        testItem.addRelationship(toolItem);
                    }

                    Item testRelItem = operationItem.createRelationship("mpp_OperationTest", "skip");
                    testRelItem.setRelatedItem(testItem);
                    //testRelItem.setNewID();
                    testRelItem.setID(cmfTestItem.Attributes["d"].Value);

                    
                }
            }


            return operationItems;
        }

        /// <summary>
        /// 动态读取QuickEditForm属性
        /// </summary>
        /// <param name="inn">innovator</param>
        /// <returns>返回属性列表字符串</returns>
        public List<string> getQuickEditFormPropertys(Innovator inn)
        {
            string[] forms = { "mpp_OperationQuickEdit", "mpp_StepQuickEdit" };
            List<string> propertys = new List<string>();

            string aml = "<AML><Item type='form' action='get' select='id'>" +
                    "<keyed_name>{0}</keyed_name>" +
                    "<Relationships>" +
                    "<Item type='Body' select='id'>" +
                        "<Relationships>" +
                            "<Item type='Field' select='propertytype_id(name)'>" +
                            "</Item>" +
                        "</Relationships>" +
                    "</Item>" +
                    "</Relationships>" +
                "</Item></AML>";

            foreach (string formName in forms)
            {
                string propertyStr = "";
                Item formItem = inn.applyAML(String.Format(aml, formName));
                Item fields = formItem.getItemsByXPath("//Item[@type='Field']");
                for (int i = 0; i < fields.getItemCount(); i++)
                {
                    Item field = fields.getItemByIndex(i);
                    Item property = field.getPropertyItem("propertytype_id");
                    if (property == null)
                    {
                        continue;
                    }
                    propertyStr += "," + property.getProperty("name");
                }
                propertys.Add(propertyStr);
            }

            return propertys;
        }

        /// <summary>
        /// 读取工时定额与检验项目页面网格数据
        /// </summary>
        /// <param name="mpp">MPP对象</param>
        /// <returns>返回网格数据</returns>
        public Item GetTestWorkHourTreeGrid(Item mpp)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            System.Web.Script.Serialization.JavaScriptSerializer JSONSerializer = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            bool isUsedPQD = CheckIsUsedPQD();

            string plan_id = mpp.getProperty("processplan_id");
            string plan_number = mpp.getProperty("processplan_number");
            string plan_name = mpp.getProperty("processplan_name");
            string location_id = mpp.getProperty("location_id", "");

            Innovator inn = mpp.getInnovator();
            int uniqueId = 0;

            StringBuilder gridStyle = new StringBuilder();
            gridStyle.Append("<?xml version='1.0' encoding='utf-8'?>");
            gridStyle.Append("<table");
            gridStyle.Append(" font='Microsoft Sans Serif-8'");
            gridStyle.Append(" sel_bgColor='steelbue'");
            gridStyle.Append(" sel_TextColor='white'");
            gridStyle.Append(" header_BgColor='buttonface'");
            gridStyle.Append(" treelines='1'");
            gridStyle.Append(" editable='true'");
            gridStyle.Append(" draw_grid='true'");
            gridStyle.Append(" multiselect='true'");
            gridStyle.Append(" column_draggable='true'");
            gridStyle.Append(" enableHtml='false'");
            gridStyle.Append(" enterAsTab='false'");
            gridStyle.Append(" bgInvert='true'");
            gridStyle.Append(" xmlns:msxsl='urn:schemas-microsoft-com:xslt'");
            gridStyle.Append(" xmlns:aras='http://www.aras.com'");
            gridStyle.Append(" xmlns:usr='urn:the-xml-files:xslt'>");
            gridStyle.Append(" <thead>");
            gridStyle.Append("  <th align='c'>Number</th>");
            gridStyle.Append("  <th align='c'>Name</th>");
            gridStyle.Append("  <th align='c'>Hours</th>");
            gridStyle.Append(" </thead>");
            gridStyle.Append(" <columns>");
            gridStyle.Append("  <column width='260' edit='NOEDIT' align='l' order='10' colname='c_item_number_mbom'/>");
            gridStyle.Append("  <column width='260' edit='NOEDIT' align='l' order='20' colname='c_name'/>");
            gridStyle.Append("  <column width='90' edit='FIELD' align='c' order='30' colname='c_bcs_hours'/>");
            gridStyle.Append(" </columns>");
            gridStyle.Append("</table>");

            //Root--ProcessPlan
            ItemJson rootPlan = new ItemJson();
            rootPlan.uniqueId = uniqueId;
            rootPlan.expanded = "true";
            rootPlan.icon = "../images/ProcessPlan.svg";
            rootPlan.expandedIcon = "../images/ProcessPlan.svg";
            ItemUserdataJson rootPlanUserdata = new ItemUserdataJson();
            rootPlanUserdata.id = plan_id;
            rootPlan.userdata = rootPlanUserdata;
            rootPlan.fields = new List<object>
            {
                Escape(plan_number),
                Escape(plan_name),
                ""
            };
            rootPlan.children = new List<ItemJson>();

            string bcs_location = !string.IsNullOrEmpty(location_id) ? "<OR><bcs_location>" + location_id + "</bcs_location><bcs_location condition='is null'></bcs_location></OR>" : string.Empty;
            string aml = $"<AML><Item type='mpp_operation' action='get' select='sort_order,name,bcs_hours'><source_id>{plan_id}</source_id>{bcs_location}" +
                (!isUsedPQD ? $"<Relationships><Item type='mpp_OperationTest' select='related_id(item_number,name)' >{bcs_location}</Item></Relationships>" : "") +
                $"<Relationships><Item type='mpp_OperationTest' select='related_id(item_number,name)' >{bcs_location}</Item></Relationships>"+
                "</Item></AML>";
            Item operations = inn.applyAML(aml);

            Dictionary<string, string> itemTypeIcons = new Dictionary<string, string>();
            if (isUsedPQD)
            {
                operations = getRelationPQDData(plan_id, operations,ref itemTypeIcons);
            }

            for (int i = 0; i < operations.getItemCount(); i++)
            {
                Item operation = operations.getItemByIndex(i);

                ItemJson operationItem = new ItemJson();
                operationItem.uniqueId = uniqueId++;
                operationItem.expanded = "true";
                operationItem.icon = "../images/ProcessOperation.svg";
                operationItem.expandedIcon = "../images/ProcessOperation.svg";

                ItemUserdataJson operationItemUserdata = new ItemUserdataJson();
                operationItemUserdata.id = operation.getID();
                operationItemUserdata.oid = operation.getID();
                operationItemUserdata.pid = plan_id;
                operationItem.userdata = operationItemUserdata;
                operationItem.fields = new List<object>
                {
                    Escape(operation.getProperty("sort_order", "")),
                    Escape(operation.getProperty("name", "")),
                    Escape(operation.getProperty("bcs_hours", ""))
                };

                operationItem.children = new List<ItemJson>();
                Item operationTests = operation.getRelationships("mpp_OperationTest");
                for (int j = 0; j < operationTests.getItemCount(); j++)
                {
                    Item operationTest = operationTests.getItemByIndex(j);
                    Item test = operationTest.getRelatedItem();
                    ItemJson operationTestItem = new ItemJson();
                    operationTestItem.uniqueId = uniqueId++;
                    operationTestItem.expanded = "true";
                    operationTestItem.icon = "../images/ViewWorkflow.svg";
                    operationTestItem.expandedIcon = "../images/ViewWorkflow.svg";

                    ItemUserdataJson operationTestItemUserdata = new ItemUserdataJson();
                    operationTestItemUserdata.id = test.getID();
                    operationTestItemUserdata.ocid = operationTest.getID();
                    operationTestItemUserdata.oid = operation.getID();
                    operationTestItem.userdata = operationTestItemUserdata;
                    operationTestItem.fields = new List<object>
                    {
                        Escape(test.getProperty("item_number", "")),
                        Escape(test.getProperty("name", "")),
                        ""
                    };

                    if (isUsedPQD)
                    {
                        operationTestItem.children = new List<ItemJson>();
                        Item testToolItems = test.getRelationships("mpp_TestTool");
                        int _toolCount = testToolItems.getItemCount();
                        for (int m = 0; m < _toolCount; m++)
                        {
                            Item testToolItem = testToolItems.getItemByIndex(m);
                            Item toolItem = testToolItem.getRelatedItem();
                            string itemTypeIcon = itemTypeIcons[toolItem.getType()];

                            ItemJson toolItemJson = new ItemJson();
                            toolItemJson.uniqueId = uniqueId++;
                            toolItemJson.expanded = "false";
                            toolItemJson.icon = itemTypeIcon;
                            toolItemJson.expandedIcon = itemTypeIcon;

                            ItemUserdataJson toolUserDataJson = new ItemUserdataJson();
                            toolUserDataJson.id = toolItem.getID();
                            toolUserDataJson.ocid = testToolItem.getID();
                            toolUserDataJson.oid = test.getID();
                            toolItemJson.userdata = toolUserDataJson;
                            toolItemJson.fields = new List<object>
                            {
                                Escape(toolItem.getProperty("item_number", "")),
                                Escape(toolItem.getProperty("name", "")),
                                ""
                            };
                            operationTestItem.children.Add(toolItemJson);
                        }
                    }

                    operationItem.children.Add(operationTestItem);
                }

                rootPlan.children.Add(operationItem);
            }

            Item result = inn.newItem("any");
            result.setProperty("mbomGridHeader", gridStyle.ToString());
            result.setProperty("mbomDataJson", JSONSerializer.Serialize(rootPlan));
            result.setProperty("uniqueid", uniqueId.ToString());
            return result;
        }

        class ItemJson
        {
            public int uniqueId;
            public string expanded;
            public string icon;
            public string expandedIcon;
            public ItemUserdataJson userdata;
            public List<object> fields;
            public List<ItemJson> children;
        }

        class ItemUserdataJson
        {
            public string id; //related PartId
            public string cpid; //ChildProcessPlanId
            public string pid; //ProcessPlanId
            public string ocid; //OpConsumedPartId
            public string oid; //OperationId
            public string eonly; //isEbomOnly
            public string comp; //isComponent
            public string bad; //isBad
            public string buy; //IsBuyNotMake
            public int gen; //related generation
            public string rev; //related revision
            public string conf; //related config_id
            public int level;
            public string ngen; //id of Part with new Generation in EBOM, is set only if status set - new Version Available
        }

        /// <summary>
        /// MPP导出至PDF
        /// </summary>
        /// <param name="mpp">MPP对象</param>
        /// <returns>返回File对象</returns>
        public Item mppExport2pdf(Item mpp,dynamic CCO)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            Innovator inn = this.innovator;
            string baseUrl = CCO.Request.Url.ToString().Replace("Server/InnovatorServer.aspx", "Client/scripts/webEditor/ueditor/TemporaryFile/");
            string baseDir = CCO.Server.MapPath("../Client/scripts/WebEditor/ueditor/TemporaryFile/");

            string html = mpp.getProperty("html");

            XmlDocument xmlDoc = inn.newXMLDocument();
            using (Sgml.SgmlReader sgmlReader = new Sgml.SgmlReader())
            {
                sgmlReader.DocType = "HTML";
                sgmlReader.InputStream = new StringReader(html);
                using (StringWriter stringWriter = new StringWriter())
                {
                    using (XmlTextWriter xmlWriter = new XmlTextWriter(stringWriter))
                    {
                        while (!sgmlReader.EOF)
                        {
                            xmlWriter.WriteNode(sgmlReader, true);
                        }
                        string xml = stringWriter.ToString().Replace("xmlns=\"http://www.w3.org/1999/xhtml\"", "");
                        xmlDoc.LoadXml(xml);
                    }
                }
            }

            XmlNodeList imgs = xmlDoc.GetElementsByTagName("img");
            List<string> fileList = new List<string>();
            foreach (XmlNode nd in imgs)
            {
                if (nd.Attributes["id"] == null) { continue; }
                Item tp_image = inn.getItemById("tp_image", nd.Attributes["id"].Value);
                if (tp_image == null) { continue; }

                string img_src = tp_image.getProperty("src", "");
                if (img_src == "" || img_src.ToLower().IndexOf("vault:///?fileid=") < 0)
                { continue; }
                string fileId = img_src.ToLower().Replace("vault:///?fileid=", "");
                Item file = inn.getItemById("File", fileId);
                if (file == null) { continue; }

                string fileName = file.getProperty("filename");
                string suffix = System.IO.Path.GetExtension(fileName);
                fileName = inn.getNewID() + suffix;
                string filePath = baseDir + fileName;

                inn.getConnection().DownloadFile(file, filePath, true);
                if (File.Exists(filePath))
                {
                    fileList.Add(filePath);
                    nd.Attributes["src"].Value = baseUrl + fileName;
                }
            }
            html = UnTransferred(xmlDoc.OuterXml);

            string newID = inn.getNewID();
            string pdfFileName = newID + ".pdf";
            string htmlFilePath = baseDir + newID + ".html";
            string pdfFilePath = baseDir + pdfFileName;

            if (File.Exists(htmlFilePath))
            {
                File.Delete(htmlFilePath);
            }
            FileStream fs1 = new FileStream(htmlFilePath, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs1);
            sw.WriteLine(html);
            sw.Close();
            fs1.Close();

            var p = new System.Diagnostics.Process();
            p.StartInfo.FileName = CCO.Server.MapPath("../Client/scripts/WebEditor/ueditor/") + "bin/wkhtmltopdf.exe";
            p.StartInfo.Arguments = "--disable-smart-shrinking \"" + htmlFilePath + "\" \"" + pdfFilePath + "\"";
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.RedirectStandardInput = true;
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.RedirectStandardError = false;
            p.StartInfo.CreateNoWindow = false;
            p.Start();
            p.WaitForExit();

            File.Delete(htmlFilePath);
            foreach (string filePath in fileList)
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }

            if (File.Exists(pdfFilePath))
            {
                Item file = inn.newItem("File", "add");
                file.setProperty("filename", pdfFileName);
                file.attachPhysicalFile(pdfFilePath);
                file = file.apply();

                File.Delete(pdfFilePath);

                return file;
            }

            return inn.newError("导出PDF失败");
        }

        /// <summary>
        /// 读取导入MPP SOP的Excel文件
        /// </summary>
        /// <param name="mpp">MPP对象</param>
        /// <returns>返回Excel转换成的HTML格式内容</returns>
        public Item readMPPTemplateFile(Item file, dynamic CCO)
        {
            if (CheckLicense() == false)
            {
                return Cinn.newError(CstrErrMessage);
            }

            Innovator inn = this.innovator;

            string fileName = file.getProperty("filename");
            //获取文件后缀
            string suffix = System.IO.Path.GetExtension(fileName).ToLower();
            if (suffix != ".xls" && suffix != ".xlsx")
            {
                return inn.newError("只可选取Excel文件!");
            }
            //文件下载路径
            string filePath = CCO.Server.MapPath("../Client/scripts/WebEditor/ueditor/TemporaryFile/") + inn.getNewID() + suffix;
            //下载文件
            inn.getConnection().DownloadFile(file, filePath, true);

            //检查文件是否下载成功
            if (!System.IO.File.Exists(filePath))
            {
                return inn.newError("下载文件到服务端失败!");
            }

            System.IO.FileStream fs = null;
            string html;
            try
            {
                using (fs = System.IO.File.OpenRead(filePath))
                {
                    NPOI.SS.UserModel.IWorkbook workbook = null;
                    if (suffix == ".xls")
                    {
                        workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(fs);
                    }
                    else
                    {
                        workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(fs);
                    }


                    NPOI.SS.Converter.ExcelToHtmlConverter excelToHtmlConverter = new NPOI.SS.Converter.ExcelToHtmlConverter();

                    // 设置输出参数
                    excelToHtmlConverter.OutputColumnHeaders = false;
                    excelToHtmlConverter.OutputHiddenColumns = true;
                    //excelToHtmlConverter.OutputHiddenRows = false;
                    //excelToHtmlConverter.OutputLeadingSpacesAsNonBreaking = true;
                    excelToHtmlConverter.OutputRowNumbers = false;
                    //excelToHtmlConverter.UseDivsToSpan = false;

                    // 处理的Excel文件
                    excelToHtmlConverter.ProcessWorkbook(workbook);
                    // fs.Close();

                    //添加表格样式
                    //excelToHtmlConverter.Document.InnerXml =excelToHtmlConverter.Document.InnerXml.Insert(
                    //    excelToHtmlConverter.Document.InnerXml.IndexOf("<head>", 0) + 6,@"<style>table, td, th{border:1px solid green;}th{background-color:green;color:white;}</style>");
                    html = excelToHtmlConverter.Document.GetElementsByTagName("table")[0].OuterXml;
                }


                //删除临时文件
                System.IO.File.Delete(filePath);

                return inn.newResult(html);
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                }
                //删除临时文件
                System.IO.File.Delete(filePath);
                return inn.newError("读取文件内容失败:" + ex.Message);
            }
        }

        #endregion

        #region "                   方法區(內部使用)"

        private bool CheckLicense()
        {
            try
            {
                if (!CbwGeneric.IsLicenseUseRuntimeFunctionByName(CstrLicenseCode))
                {
                    CstrErrMessage = CbwGeneric.ErrorException;

                    if (CstrErrMessage == "") CstrErrMessage = CoCommon.GetMessageByMsgId("msg_gen_000024", "授權碼驗證不正確");

                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                CstrErrMessage = ex.Message;
                return false;
            }
        }

        private string Escape(string data)
        {
            return System.Security.SecurityElement.Escape(data);
        }

        private string UnTransferred(string Meaning)
        {
            //转义字符变换成普通字符
            Meaning = Meaning.Replace("&lt;", "<");
            Meaning = Meaning.Replace("&gt;", ">");
            Meaning = Meaning.Replace("&apos;", "'");
            Meaning = Meaning.Replace("&quot;", "\"");
            Meaning = Meaning.Replace("&amp;", "&");
            return Meaning;
        }

        /// <summary>
        /// 检查是否有使用PQD联动
        /// </summary>
        /// <returns>true:有使用 false:未使用</returns>
        private bool CheckIsUsedPQD()
        {
            return !innovator.applyMethod("bcs_MPP_CheckPQDIsUsed", "").isError();
        }

        #endregion
    }
}
