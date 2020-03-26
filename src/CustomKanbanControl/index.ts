import {IInputs, IOutputs} from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
import { format } from "path";
import { Debugger } from "inspector";
import { debug } from "util";
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class CustomKanbanControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	/* PARAMETERS */
	private _config: ConfigObject;

	private _entity: string;
	private _viewId: string;	

	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _loader: HTMLDivElement;
	
	private _customText: CustomText;
	private _userCurrency: Currency;

	private dataSet:ComponentFramework.PropertyTypes.DataSet;
	private allGroups: Array<GroupObject>;
	private allRecordStages: Array<RecordStageObject>;
	private allCurrencies: Array<Currency>;
	private dataList:Array<DataObject>;
	private allMetadata:Array<OptionSetMeta>;
	private entityMeta:EntityMeta;
	private attributesMeta:Array<AttributeMeta>;
	private editedRecordId: string;
	private dataLoaded: boolean;
	private stagesLoaded: boolean;
	private currenciesLoaded: boolean;
	private entitiesLoaded: boolean;
	private attributesMetadataLoaded: boolean;
	private metadataLoaded: boolean;
	private metadata1Loaded: boolean;
	private metadata2Loaded: boolean;
	private metadata3Loaded: boolean;
	private metadata4Loaded: boolean;

	private _sysColumns: Array<string>;

	/*
	*	Events
	*/
	private _dragStart: EventListenerOrEventListenerObject;
	private _dragEnd: EventListenerOrEventListenerObject;
	private _dragEnter: EventListenerOrEventListenerObject;
	private _dragLeave: EventListenerOrEventListenerObject;
	private _dragOver: EventListenerOrEventListenerObject;
	private _dropped: EventListenerOrEventListenerObject;
	private _draggedObject: DataObject;

	private _valueClicked: EventListenerOrEventListenerObject;
	private _saveElement: EventListenerOrEventListenerObject;
	private _linkClicked: EventListenerOrEventListenerObject;
	private _moreClicked: EventListenerOrEventListenerObject;

	/**
	 * Empty constructor.
	 */
	constructor()
	{
		this._customText = { Save: "Save", More: "More records", Empty: "Empty" };
		this.dataLoaded = false;
		this.stagesLoaded = false;
		this.entitiesLoaded = false;
		this.currenciesLoaded = false;
		this.attributesMetadataLoaded = false;
		this.metadataLoaded = false;
		this.metadata1Loaded = false;
		this.metadata2Loaded = false;
		this.metadata3Loaded = false;
		this.metadata4Loaded = false;
	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		//@ts-ignore
		this._entity = context.page.entityTypeName;
		this._viewId = context.parameters.sampleDataSet.getViewId();

		var language = context.userSettings.languageId;
		
		//var cultureName = context.userSettings.for;

		this._context = context;
		this._container = container;
		this._container.setAttribute("class", "pcfContainer");
		this._container.setAttribute("class", "mainContainer");

		this._dragStart = this.dragStart.bind(this);
		this._dragEnd = this.dragEnd.bind(this);
		this._dragEnter = this.dragEnter.bind(this);
		this._dragLeave = this.dragLeave.bind(this);
		this._dragOver = this.dragOver.bind(this);
		this._dropped = this.dropped.bind(this);

		this._valueClicked = this.valueClicked.bind(this);
		this._saveElement = this.saveElement.bind(this);
		this._linkClicked = this.linkClicked.bind(this);
		this._moreClicked = this.moreClicked.bind(this);

		this.createLoader();
		var self = this;

		this.getConfig(context, function(ok:boolean, data:ConfigObject, error:any){
			if(ok){
				self._config = data;

				var allColumns = context.parameters.sampleDataSet.columns.map(a => a.name); 
				// if(allColumns.indexOf(data.amountAttribute) == -1) {
				// 	self.showError("Amount column is not defined in view (" + data.amountAttribute + ")");
				// 	return;
				// }
				// if(allColumns.indexOf(data.nameAttribute) == -1) {
				// 	self.showError("Name column is not defined in view (" + data.nameAttribute + ")");
				// 	return;
				// }
				// if(data.groupType == GroupType.Entity && allColumns.indexOf(data.entityAttribute) == -1) {
				// 	self.showError("Entity Attribute column is not defined in view (" + data.entityAttribute + ")");
				// 	return;
				// }
				
				// if(data.groupType == GroupType.Picklist && allColumns.indexOf(data.picklistAttribute) == -1) {
				// 	self.showError("Picklist Attribute column is not defined in view (" + data.picklistAttribute + ")");
				// 	return;
				// }

				if(data.groupType == GroupType.Picklist && data.picklistAttribute == "statecode") {
					self.showError("StateCode is not supported for Picklist Attribute, please use StatusCode instead");
					return;
				}

				self._sysColumns = [];

				var refresh = false;
				if(allColumns.indexOf(data.amountAttribute) == -1) {
					if(self._context.parameters.sampleDataSet.addColumn){
						self._context.parameters.sampleDataSet.addColumn(data.amountAttribute);
						refresh = true;
					}else{
						self.showError("Amount column is not defined in view (" + data.amountAttribute + ")");
						return;
					}
				}
				if(allColumns.indexOf("transactioncurrencyid") == -1) {
					if(self._context.parameters.sampleDataSet.addColumn){
						self._context.parameters.sampleDataSet.addColumn("transactioncurrencyid");
						self._sysColumns.push("transactioncurrencyid");
						refresh = true;
					}else{
						self.showError("TransactionCurrencyId column is not defined in view");
						return;
					}
				}
				
				if(allColumns.indexOf(data.nameAttribute) == -1) {
					if(self._context.parameters.sampleDataSet.addColumn){
						self._context.parameters.sampleDataSet.addColumn(data.nameAttribute);
						refresh = true;
					}else{
						self.showError("Name column is not defined in view (" + data.nameAttribute + ")");
						return;
					}
				}
				if(data.groupType == GroupType.Entity && allColumns.indexOf(data.entityAttribute) == -1) {
					if(self._context.parameters.sampleDataSet.addColumn){
						self._context.parameters.sampleDataSet.addColumn(data.entityAttribute);
						self._sysColumns.push(data.entityAttribute);

						refresh = true;
					}else{
						self.showError("Entity Attribute column is not defined in view (" + data.entityAttribute + ")");
						return;
					}
				}
				
				if(data.groupType == GroupType.Picklist && allColumns.indexOf(data.picklistAttribute) == -1) {
					if(self._context.parameters.sampleDataSet.addColumn){
						self._context.parameters.sampleDataSet.addColumn(data.picklistAttribute);
						self._sysColumns.push(data.picklistAttribute);
						refresh = true;
					}else{
						self.showError("Picklist Attribute column is not defined in view (" + data.picklistAttribute + ")");
						return;
					}
				}

				
				if(refresh){
					self._context.parameters.sampleDataSet.refresh();
				}

				
				self.getCurrencies(context, function(ok: boolean, data: Array<Currency>, error:string){
					if(ok){
						self.allCurrencies = data;
						self.currenciesLoaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});

				if(self._config.groupType == GroupType.BPF){
					self.getProcessInfos(context, function(infosOk: boolean, infosData:Array<string>, error:any){
						if(infosOk){
							var infos = infosData;
							self.getProcessStages(context, function(ok:boolean, data:Array<GroupObject>, error:any){
								if(ok){
									data.forEach(a => a.order = infos.indexOf(a.id));
									self.allGroups = data.sort(function(a,b){ return a.order > b.order ? 1 : -1; });
									self.stagesLoaded = true;
									self.createLayout();
								}
							});
						}else{
							debugger;
						}
					});
				}else if(self._config.groupType == GroupType.Entity){
					self.getEntities(context, function(ok:boolean, data:Array<GroupObject>, error:any){
						if(ok){
							// self.allGroups = data.sort(function(a,b){ return a.order > b.order ? 1 : -1; });
							self.entitiesLoaded = true;
							self.allGroups = data;
							self.createLayout();
						}else{
							debugger;
						}
					});
				}
				
				self.getEntityMeta(self._entity, context, function(ok:boolean, data: EntityMeta, error:string){
					if(ok){
						self.entityMeta = data;
						self.metadataLoaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});
				self.getEntityAttributesMeta(self._entity, context, function(ok:boolean, data: Array<AttributeMeta>, error:string){
					if(ok){
						self.attributesMeta = data;
						self.attributesMetadataLoaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});
				
				self.allMetadata = [];
				self.getEntityOptions("Picklist", context, function(ok:boolean, data:Array<OptionSetMeta>, error:string){
					if(ok){
						self.allMetadata = self.allMetadata.concat(data);
						self.metadata1Loaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});
				self.getEntityOptions("State", context, function(ok:boolean, data:Array<OptionSetMeta>, error:string){
					if(ok){
						self.allMetadata = self.allMetadata.concat(data);
						self.metadata2Loaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});
				self.getEntityOptions("Status", context, function(ok:boolean, data:Array<OptionSetMeta>, error:string){
					if(ok){
						self.allMetadata = self.allMetadata.concat(data);
						self.metadata3Loaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});
				self.getEntityOptions("Boolean", context, function(ok:boolean, data:Array<OptionSetMeta>, error:string){
					if(ok){
						self.allMetadata = self.allMetadata.concat(data);
						self.metadata4Loaded = true;
						self.createLayout();
					}else{
						debugger;
					}
				});
			}else{
				self.showError(error);
			}
		});

		
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		if (!context.parameters.sampleDataSet.loading) {
			if (context.parameters.sampleDataSet.paging != null && context.parameters.sampleDataSet.paging.hasNextPage == true) {
				//set page size
				context.parameters.sampleDataSet.paging.setPageSize(5000);
				//load next paging
				context.parameters.sampleDataSet.paging.loadNextPage();
			} else {
				this.dataLoaded = true;
				this._context = context;
				this.dataSet = context.parameters.sampleDataSet;

				var self = this;
				if(this._config.groupType == GroupType.BPF){
					this.getProcessRecordsStages(context, function(ok: boolean, result:Array<RecordStageObject>, error:any){
						if(ok){
							self.allRecordStages = result;
							self.createLayout();
						}else{
							debugger;
						}
					});
				}else{
					this.createLayout();
				}
			}
		}
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}




	private dragStart(evt:Event){
		var element = evt.target as HTMLDivElement;
		var current = this.dataList.filter(a => a.DivElement == element);
		if(current.length > 0){
			var data = current[0];
			this.allGroups.filter(s => s.id != data.groupid).forEach(stage => {
				stage.DropContainer?.classList.add("dropContainerVisible");
			});

			this._draggedObject = data;
		}
	}
	private dragEnd(evt:Event){
		this.allGroups.forEach(stage => {
			stage.DropContainer?.classList.remove("dropContainerVisible");
		});
	}
	private dragEnter(evt:Event){
		var element = evt.target as HTMLDivElement;
		element.classList.add("dropContainerOver");
		var idx = parseInt(element.getAttribute("elementIdx")?? "0");
		element.setAttribute("style", "border-color: " + this.getColor(idx, "ff") + "; background: " + this.getColor(idx, "30"));
	}
	private dragLeave(evt:Event){
		var element = evt.target as HTMLDivElement;
		element.classList.remove("dropContainerOver");
		element.setAttribute("style", "");
	}
	private dragOver(evt:Event){
		var e = evt as DragEvent;

		if (evt.preventDefault) {
			evt.preventDefault(); // Necessary. Allows us to drop.
		}
		var data = e.dataTransfer;
		if(data) data.dropEffect = "move";
		return false;
	}
	private dropped(evt:Event){
		if (evt.stopPropagation) {
			evt.stopPropagation(); // Stops some browsers from redirecting.
		}
		
		var targetGroup = this.allGroups.filter(a => a.DropContainer == evt.target);
		if(targetGroup.length > 0){
			this.showLoader(true);
			var self = this;

			var entity = this._entity;
			var upd = {};
			var id = "";

			if(this._config.groupType == GroupType.BPF){
				entity = this._config.bpfEntity;
				var filtered = this.allRecordStages.filter(a => a.recordid == this._draggedObject.id);
				if(filtered.length > 0){
					id = filtered[0].id;
					//@ts-ignore
					upd["activestageid@odata.bind"] = "/processstages(" + targetGroup[0].id + ")";
				}else{
					//@ts-ignore
					upd["activestageid@odata.bind"] = "/processstages(" + targetGroup[0].id + ")";
					//@ts-ignore
					upd[this._config.bpfRecordAttribute +"@odata.bind"] = "/" + this.entityMeta.LogicalCollectionName + "(" + this._draggedObject.id + ")";

					this._context.webAPI.createRecord(entity, upd).then(
						function(){
							self._context.parameters.sampleDataSet.refresh();
						},function(errorResponse:any){
							debugger;
							self.showLoader(false);
							self.showError(errorResponse);
						}
					)

					return;
				}
			}
			else if(this._config.groupType == GroupType.Picklist){
				//TODO: Update picklist and manage statuscode (update statecode too)
				id = this._draggedObject.id;

				if(this._config.picklistAttribute == "statuscode"){
					var opts = this.allMetadata.filter(a => a.logicalname == "statuscode")[0].options;
					var opt = opts.filter(a => a.value == parseInt(targetGroup[0].id))[0];

					//@ts-ignore
					upd["statecode"] = opt.parentValue;
					//@ts-ignore
					upd["statuscode"] = opt.value;
				}
				else{
					//@ts-ignore
					upd[this._config.picklistAttribute] = parseInt(targetGroup[0].id);
				}
			}
			else if(this._config.groupType == GroupType.Entity){
				id = this._draggedObject.id;
				//@ts-ignore
				upd[this._config.entityAttribute + "@odata.bind"] = "/" + this._config.entityLogicalName + "(" + targetGroup[0].id + ")";
			}
	


			this._context.webAPI.updateRecord(entity, id, upd).then(
				function(){
					self._context.parameters.sampleDataSet.refresh();
				},function(errorResponse:any){
					debugger;
					self.showLoader(false);
					self.showError(errorResponse);
				}
			)

		}
		
		
		return false;
	}


	
	private getConfig(context: ComponentFramework.Context<IInputs>, callback: Function){
		var entity = "dem_customkanbanconfig";
		var options = "?$select=dem_amountattribute,dem_entityicon,dem_showemptygroup,dem_maxlines,dem_picklistattribute,dem_entityattribute,dem_entitylogicalname,dem_entityfetchxml,dem_entityidattribute,dem_entitynameattribute,dem_bpfrecordattribute,dem_sectionminwidth,dem_colorfullsections,dem_groupbytype,dem_bpfentity,_dem_amountcurrency_value,_dem_businessprocessflowid_value,dem_colorslist,dem_editableattributes,dem_nameattribute,dem_showamount,dem_viewid&$filter=dem_viewid eq '" + this._viewId + "'";
		var self = this;
		context.webAPI.retrieveMultipleRecords(entity, options).then(
			function(response:ComponentFramework.WebApi.RetrieveMultipleResponse){
				var entities = response.entities;
				var conf:ConfigObject;
				if(entities.length > 0){
					var amountattribute = entities[0]["dem_amountattribute"];
					var nameattribute = entities[0]["dem_nameattribute"];
					var processid = entities[0]["_dem_businessprocessflowid_value"];
					var colorslist = entities[0]["dem_colorslist"];
					var editableattributes = entities[0]["dem_editableattributes"];
					var showamount = entities[0]["dem_showamount"];
					var colorsections = entities[0]["dem_colorfullsections"];
					var sectionminwidth = entities[0]["dem_sectionminwidth"];
					var groupbytype = entities[0]["dem_groupbytype"];
					var bpfentity = entities[0]["dem_bpfentity"];
					var bpfrecordattribute = entities[0]["dem_bpfrecordattribute"];
					var picklistattribute = entities[0]["dem_picklistattribute"];
					var entityattribute = entities[0]["dem_entityattribute"];
					var entityfetchxml = entities[0]["dem_entityfetchxml"];
					var entitylogicalname = entities[0]["dem_entitylogicalname"];
					var entityidattribute = entities[0]["dem_entityidattribute"];
					var entityNameAttribute = entities[0]["dem_entitynameattribute"];
					var maxLines = entities[0]["dem_maxlines"] ?? 10;
					var showEmptyGroup = !!entities[0]["dem_showemptygroup"];
					var entityIcon = entities[0]["dem_entityicon"];
					var currency = entities[0]["_dem_amountcurrency_value"];
					
					conf = {
						amountAttribute: amountattribute,
						nameAttribute: nameattribute,
						showAmount: showamount ? showamount : false,
						editableAttributes: editableattributes ? editableattributes.split(",") : [],
						colorsList: colorslist ? colorslist.split(",") : "#ababab",
						colorfullsections: colorsections ? colorsections : false,
						minWidthSection: sectionminwidth ? sectionminwidth : 250,
						groupType: groupbytype,
						processId: processid,
						bpfEntity: bpfentity,
						bpfRecordAttribute: bpfrecordattribute,
						picklistAttribute: picklistattribute,
						entityAttribute: entityattribute,
						entityFetchXml: entityfetchxml,
						entityLogicalName: entitylogicalname,
						entityIdAttribute: entityidattribute,
						entityNameAttribute: entityNameAttribute,
						maxRecords: maxLines,
						ShowEmptyGroup: showEmptyGroup,
						EntityIcon: entityIcon,
						currencyId: currency
					}


					if(!amountattribute) callback(false, conf, "Amount Attribute is not defined in configuration");
					if(!currency) callback(false, conf, "Currency is not defined in configuration");
					else if(!nameattribute) callback(false, conf, "Name Attribute is not defined in configuration");
					else if(!groupbytype) callback(false, conf, "GroupBy Type is not defined in configuration");
					else if(groupbytype == GroupType.BPF && !processid) callback(false, conf, "BPF id is not defined in configuration");
					else if(groupbytype == GroupType.BPF && !bpfentity) callback(false, conf, "BPF entity is not defined in configuration");
					else if(groupbytype == GroupType.BPF && !bpfrecordattribute) callback(false, conf, "BPF Record Attribute is not defined in configuration");
					else if(groupbytype == GroupType.Picklist && !picklistattribute) callback(false, conf, "Picklist Attribute is not defined in configuration");
					else if(groupbytype == GroupType.Entity && !entityattribute) callback(false, conf, "Entity Attribute is not defined in configuration");
					else if(groupbytype == GroupType.Entity && !entityfetchxml) callback(false, conf, "Entity FetchXml is not defined in configuration");
					else if(groupbytype == GroupType.Entity && !entitylogicalname) callback(false, conf, "Entity Logicalname is not defined in configuration");
					else if(groupbytype == GroupType.Entity && !entityidattribute) callback(false, conf, "Entity Id Attribute is not defined in configuration");
					else if(groupbytype == GroupType.Entity && !entityNameAttribute) callback(false, conf, "Entity Name Attribute is not defined in configuration");
					else{
						callback(true, conf);
					}
				}else{
					callback(false, undefined, "No configuration defined for this view");
				}
			},
			function(errorResponse:any){
				debugger;
				callback(false, undefined, errorResponse);
			});
	}
	private getEntityMeta(type:string, context: ComponentFramework.Context<IInputs>, callback:Function):void{
		//@ts-ignore
		var serverUrl = Xrm.Page.context.getClientUrl();
		//@ts-ignore
		var version:string = Xrm.Page.context.getVersion();
		var v = version.split(".");
		var apiVersion = [v[0],v[1]].join(".");

		var request = serverUrl + "/api/data/v" + apiVersion + "/EntityDefinitions(LogicalName='" + this._entity + "')?$select=LogicalCollectionName";
	
		var self = this;
		let req = new XMLHttpRequest();
		req.open("GET", request, true);
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.onreadystatechange = () => {
			if (req.readyState == 4 /* complete */) {
				req.onreadystatechange = null; /* avoid potential memory leak issues.*/

				if (req.status == 200) {
					var data = JSON.parse(req.response);
					var metadata: EntityMeta;

					metadata = {
						LogicalName: type,
						LogicalCollectionName: data.LogicalCollectionName
					}
					
					callback(true, metadata);
				} else {
					var error = JSON.parse(req.response).error;
					callback(false, undefined, error);
				}
			}
		};
		req.send();
	}
	private getEntityAttributesMeta(type:string, context: ComponentFramework.Context<IInputs>, callback:Function):void{
		//@ts-ignore
		var serverUrl = Xrm.Page.context.getClientUrl();
		//@ts-ignore
		var version:string = Xrm.Page.context.getVersion();
		var v = version.split(".");
		var apiVersion = [v[0],v[1]].join(".");

		var request = serverUrl + "/api/data/v" + apiVersion + "/EntityDefinitions(LogicalName='" + this._entity + "')/Attributes?$select=LogicalName,AttributeType";
	
		var self = this;
		let req = new XMLHttpRequest();
		req.open("GET", request, true);
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.onreadystatechange = () => {
			if (req.readyState == 4 /* complete */) {
				req.onreadystatechange = null; /* avoid potential memory leak issues.*/

				if (req.status == 200) {
					var data = JSON.parse(req.response).value as Array<any>;
					var metadata: Array<AttributeMeta>;
					metadata = [];

					for(var i = 0; i < data.length; i++){
						metadata.push({
							//
							LogicalName: data[i].LogicalName,
							Type: data[i].AttributeType
						})
					}
					
					callback(true, metadata);
				} else {
					var error = JSON.parse(req.response).error;
					callback(false, undefined, error);
				}
			}
		};
		req.send();
	}
	private getEntityOptions(type: string, context: ComponentFramework.Context<IInputs>, callback:Function):void{
		//@ts-ignore
		var serverUrl = Xrm.Page.context.getClientUrl();
		//@ts-ignore
		var version:string = Xrm.Page.context.getVersion();
		var v = version.split(".");
		var apiVersion = [v[0],v[1]].join(".");

		var request = serverUrl + "/api/data/v" + apiVersion + "/EntityDefinitions(LogicalName='" + this._entity + "')/Attributes/Microsoft.Dynamics.CRM." + type + "AttributeMetadata?$select=LogicalName,DisplayName&$expand=OptionSet,GlobalOptionSet";
	
		var self = this;
		let req = new XMLHttpRequest();
		req.open("GET", request, true);
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.onreadystatechange = () => {
			if (req.readyState == 4 /* complete */) {
				req.onreadystatechange = null; /* avoid potential memory leak issues.*/

				if (req.status == 200) {
					var options = JSON.parse(req.response).value as Array<any>;
					var list:Array<OptionSetMeta> = []; 
					options.forEach(element => {
						var values:Array<OptionMeta> = [];

						if(type == "Boolean"){
							values.push({
								value: element.GlobalOptionSet.TrueOption.Value as number,
								label: element.GlobalOptionSet.TrueOption.Label.UserLocalizedLabel.Label as string,
								parentValue: undefined
							});
							values.push({
								value: element.GlobalOptionSet.FalseOption.Value as number,
								label: element.GlobalOptionSet.FalseOption.Label.UserLocalizedLabel.Label as string,
								parentValue: undefined
							});
						}else{

							(element.GlobalOptionSet.Options as Array<any>).forEach(o => {
								values.push({
									value: o.Value as number,
									label: o.Label.UserLocalizedLabel.Label as string,
									parentValue: (type == "Status" ? o.State : undefined)
								});
							});
						}

						list.push({
							logicalname: element.LogicalName,
							type: type,
							options: values
						});
					});
					callback(true, list);
				} else {
					var error = JSON.parse(req.response).error;
					callback(false, undefined, error);
				}
			}
		};
		req.send();
	}
	
	private extractProcessSteps(steps:any):Array<string>{
		var list = [];
		if(steps.hasOwnProperty("list")){
			for(var i = 0; i < steps.list.length; i++){
				var t = steps.list[i];
				if(t.hasOwnProperty("stageId")){
					list.push(t.stageId);
				}
				if(t.steps) list = list.concat(this.extractProcessSteps(t.steps));
			}
		}
		return list;
	}
	private getProcessInfos(context: ComponentFramework.Context<IInputs>, callback:Function){
		var entity = "workflow";
		var options = "?$select=clientdata&$filter=workflowid eq " + this._config.processId;
		var self = this;
		context.webAPI.retrieveMultipleRecords(entity, options).then(
			function(response:ComponentFramework.WebApi.RetrieveMultipleResponse){
				if(response.entities.length > 0){
					var data = JSON.parse(response.entities[0]["clientdata"]);
					var list = self.extractProcessSteps(data.steps);
					callback(true, list);
				}
				else{
					callback(false, [], "Process not found");
				}
			},
			function(errorResponse:any){
				debugger;
				callback(false, [], errorResponse);
			});
	}
	private getProcessStages(context: ComponentFramework.Context<IInputs>, callback:Function):void
	{
		var entity = "processstage";
		var options = "?$select=stagename&$filter=_processid_value eq "+this._config.processId;
		context.webAPI.retrieveMultipleRecords(entity, options).then(
			function(response:ComponentFramework.WebApi.RetrieveMultipleResponse){
				var entities = response.entities;
				var s:GroupObject;
				var stages = [];
				for (var i = 0; i < entities.length; i++) {
					var stagename = entities[i]["stagename"];
					s = {
						order: -1,
						name: stagename,
						id: entities[i].processstageid,
						DropContainer: undefined
					};
					stages.push(s);
				}

				callback(true, stages);
			},
			function(errorResponse:any){
				debugger;
				callback(false, [], errorResponse);
			});
	}
	private getProcessRecordsStages(context: ComponentFramework.Context<IInputs>, callback: Function):void{
		var entity = this._config.bpfEntity;
		var options = "?$select=_activestageid_value,_" + this._config.bpfRecordAttribute + "_value,businessprocessflowinstanceid&$filter=statecode eq 0";
		var self = this;
		context.webAPI.retrieveMultipleRecords(entity, options).then(
			function(response:ComponentFramework.WebApi.RetrieveMultipleResponse){
				var entities = response.entities;
				var s:RecordStageObject;
				var recordsStages = [];
				for (var i = 0; i < entities.length; i++) {
					var stageid = entities[i]["_activestageid_value"];
					var recordid = entities[i]["_" + self._config.bpfRecordAttribute + "_value"];
					var id = entities[i]["businessprocessflowinstanceid"];
					s = {
						stageid: stageid,
						recordid: recordid,
						id: id
					};
					recordsStages.push(s);
				}

				callback(true, recordsStages);
			},
			function(errorResponse:any){
				debugger;
				callback(false, [], errorResponse);
			});
	}
	private getEntities(context: ComponentFramework.Context<IInputs>, callback:Function):void
	{
		//@ts-ignore
		var serverUrl = Xrm.Page.context.getClientUrl();
		//@ts-ignore
		var version:string = Xrm.Page.context.getVersion();
		var v = version.split(".");
		var apiVersion = [v[0],v[1]].join(".");

		var request = serverUrl + "/api/data/v" + apiVersion + "/" + this._config.entityLogicalName + "?fetchXml=" + encodeURI(this._config.entityFetchXml);
	
		let req = new XMLHttpRequest();
		req.open("GET", request, true);
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.onreadystatechange = () => {
			if (req.readyState == 4 /* complete */) {
				req.onreadystatechange = null; /* avoid potential memory leak issues.*/

				if (req.status == 200) {
					var entities = JSON.parse(req.response).value as Array<any>;
					var list:Array<GroupObject> = []; 
					var idx = 0;
					entities.forEach(element => {
						list.push({
							DropContainer: undefined,
							order: idx,
							id: element[this._config.entityIdAttribute],
							name: element[this._config.entityNameAttribute] 
						});
						idx++;
					});
					callback(true, list);
				} else {
					debugger;
					var error = JSON.parse(req.response).error;
					callback(false, undefined, error);
				}
			}
		};
		req.send();
	}
	private getCurrencies(context: ComponentFramework.Context<IInputs>, callback:Function):void{
		var entity = "transactioncurrency";
		var options = "?$select=transactioncurrencyid,exchangerate,currencysymbol";
		context.webAPI.retrieveMultipleRecords(entity, options).then(
			function(response:ComponentFramework.WebApi.RetrieveMultipleResponse){
				var entities = response.entities;
				var s:Currency;
				var currencies = [];
				for (var i = 0; i < entities.length; i++) {
					s = {
						id: entities[i].transactioncurrencyid,
						rate: entities[i].exchangerate,
						symbol: entities[i].currencysymbol,
					};
					currencies.push(s);
				}

				callback(true, currencies);
			},
			function(errorResponse:any){
				debugger;
				callback(false, [], errorResponse);
			});
	}

	private createLayout(){
		if(this.dataLoaded && this.currenciesLoaded && this.metadataLoaded && this.attributesMetadataLoaded && this.metadata1Loaded && this.metadata2Loaded && this.metadata3Loaded && this.metadata4Loaded){
			
			if(this._config.groupType == GroupType.BPF && !this.stagesLoaded) return;
			if(this._config.groupType == GroupType.Entity && !this.entitiesLoaded) return;
			
			var self = this;
			self._container.innerHTML = "";
			self.createLoader();

			let amountColumn:DataSetInterfaces.Column;
			let nameColumn:DataSetInterfaces.Column;

			if(this._config.groupType == GroupType.Picklist){
				var pickAttribute = this.allMetadata.filter(a => a.logicalname == this._config.picklistAttribute);
				if(pickAttribute.length > 0){
					this.allGroups = pickAttribute[0].options.map(a => {
						return {
							order: -1,
							name: a.label,
							id: a.value.toString(),
							DropContainer: undefined
						};
					})
				}
			}

			self.dataSet.columns.forEach((column: DataSetInterfaces.Column) => {
				if(column.name == self._config.amountAttribute) amountColumn = column;
				if(column.name == self._config.nameAttribute) nameColumn = column;
			});

			if(this._config.ShowEmptyGroup){
				if(this.allGroups.filter(a => a.id == "").length == 0){
					this.allGroups.unshift({
						order: -1,
						name: this._customText.Empty,
						id: "",
						DropContainer: undefined
					});
				}
			}

			var columnsRepartition = "";
			for(var i = 0; i < this.allGroups.length; i++){
				columnsRepartition += " " + (parseFloat("100") / this.allGroups.length) + "%";
			}

			self.dataList = [];

			var fCurr = self.allCurrencies.filter(a => a.id == self._config.currencyId);
			self._userCurrency = fCurr[0];

			var data:DataObject;
			self.dataSet.sortedRecordIds.forEach((recordId: string) => {
				let currentRecord = self.dataSet.records[recordId];

				var groupid = "";
				
				if(this._config.groupType == GroupType.BPF){
					groupid = currentRecord.getFormattedValue("stageid");
					var filtered = self.allRecordStages.filter(a => a.recordid == recordId);
					if(filtered.length > 0){
						groupid = filtered[0].stageid;
					}
				}
				else if(this._config.groupType == GroupType.Picklist){
					var val = currentRecord.getValue(this._config.picklistAttribute);
					groupid = val ? val.toString() : "";
				}
				else if(this._config.groupType == GroupType.Entity){
					var val = currentRecord.getValue(this._config.entityAttribute);
					//@ts-ignore
					groupid = val ? val.id.guid.toString() : "";
				}

				var amount = amountColumn ? (currentRecord.getValue(amountColumn.name) as number ?? 0) : 0;
				if(amount > 0){
					//@ts-ignore
					var currencyId = (currentRecord.getValue("transactioncurrencyid") ? currentRecord.getValue("transactioncurrencyid").id.guid : "");
					var currency = self.allCurrencies.filter(a => a.id == currencyId);
					if(currency.length > 0){
						var temp = (amount / currency[0].rate)*self._userCurrency.rate;
						amount = temp;
					}
				}


				data = {
					id: currentRecord.getRecordId(),
					name: nameColumn ? (currentRecord.getFormattedValue(nameColumn.name) ?? "") : "",
					amount: amount,
					groupid: groupid ?? "",
					columns: self.dataSet.columns,
					entity: currentRecord,
					DivElement: undefined
				};
				self.dataList.push(data);
			});

			var body = document.createElement("div");
			body.setAttribute("class", "section");
			body.setAttribute("style", "grid-template-columns: " + columnsRepartition);

			for(var i = 0; i < this.allGroups.length; i++){
				var s = this.allGroups[i];
				//Titre de la section
				var sum = 0;
				var allCount = 0;
				var count = 0;
				
				var elements:Array<HTMLDivElement> = [];
				var records = self.dataList.filter(a => a.groupid == s.id);
				allCount = records.length;
				var createMoreButton = false;
				for(var j = 0; j < records.length; j++){
					sum += records[j].amount;
					records[j].DivElement = self.createElementHtml(records[j]);
					
					if(j < this._config.maxRecords){
						count ++;
						elements.push(records[j].DivElement as HTMLDivElement);
					}else{
						createMoreButton = true;
					}
				}

				var section = document.createElement("div");
				section.setAttribute("class", "sectionBody");
				section.setAttribute("groupid", s.id);
				section.setAttribute("max", this._config.maxRecords.toString());

				var sectionStyle = "min-width:" + self._config.minWidthSection + "px ;grid-column: "+(i + 1)+"; grid-row: 1; border-top: solid 3px " + self.getColor(i, "ff");
				if(self._config.colorfullsections){
					sectionStyle += ";background:" + self.getColor(i, "33");
				}

				section.setAttribute("style", sectionStyle);
				
				if(s.id){
					var elementContainerDrop = document.createElement("div");
					elementContainerDrop.setAttribute("class", "elementContainerDrop");
					elementContainerDrop.setAttribute("elementIdx", i.toString());
	
					section.appendChild(elementContainerDrop);
					s.DropContainer = elementContainerDrop;

					elementContainerDrop.addEventListener("dragenter", self._dragEnter);
					elementContainerDrop.addEventListener("dragleave", self._dragLeave);
					elementContainerDrop.addEventListener("dragover", self._dragOver);
					elementContainerDrop.addEventListener("drop", self._dropped);
				}
				

				

				var sectionHeader = document.createElement("div");
				sectionHeader.setAttribute("class", "sectionHeader");
				sectionHeader.innerHTML = s.name;
				if(self._config.colorfullsections){
					sectionHeader.setAttribute("style", "color:" + self.getColor(i, "ff"));
				}
				section.appendChild(sectionHeader);

				var sectionSubHeader = document.createElement("div");
				sectionSubHeader.setAttribute("class", "sectionSubHeader");
				if(self._config.colorfullsections){
					sectionSubHeader.setAttribute("style", "color:" + self.getColor(i, "ff"));
				}

				var table = document.createElement("table");
				table.setAttribute("style", "table-layout:fixed; margin: 0 5px 0 0px; width:100%");
				sectionSubHeader.appendChild(table);

				var tr = document.createElement("tr");
				var td1 = document.createElement("td");
				td1.setAttribute("style", "width:50%; text-align: left");
				td1.innerText = self.formatNumber(sum) + " " + self._userCurrency.symbol;
				
				var td2 = document.createElement("td");
				td2.setAttribute("style", "width:50%; text-align: right");
				td2.setAttribute("counter", "true");
				td2.innerText = self.formatNumber(count) + " / " + self.formatNumber(allCount);

				tr.appendChild(td1);
				tr.appendChild(td2);
				table.appendChild(tr);
				
				section.appendChild(sectionSubHeader);
				
				var elementsContainer = document.createElement("div");
				elementsContainer.setAttribute("class", "elementsContainer");


				elements.forEach(el => elementsContainer.appendChild(el));

				if(createMoreButton){

					elementsContainer.appendChild(self.createMoreButton(s.id));
				}

				section.appendChild(elementsContainer);

				body.appendChild(section);
			}

			self._container.appendChild(body);

			self.showLoader(false);
			
		}
	}
	private createMoreButton(groupid:string): HTMLButtonElement{
		var more = document.createElement("button");
		more.setAttribute("class", "moreButton");
		more.setAttribute("groupid", groupid);
		more.addEventListener("click", this._moreClicked, true);
		more.innerText = this._customText.More;

		return more;
	}
	private createElementHtml(data: DataObject):HTMLDivElement{
		var element = document.createElement("div");
		element.setAttribute("class", "element noselect");
		element.setAttribute("draggable", "true");
		element.setAttribute("recordid", data.id);
		
		var elementHeader = document.createElement("div");
		elementHeader.setAttribute("class", "elementHeader");
		element.appendChild(elementHeader);

		var elementImg = document.createElement("div");
		elementImg.setAttribute("class", "elementImg");
		if(this._config.EntityIcon){
			elementImg.setAttribute("style", "background-image: url(" + this._config.EntityIcon + ")");
		}
		elementHeader.appendChild(elementImg);

		var elementName = document.createElement("div");
		elementName.setAttribute("class", "elementName");
		elementHeader.appendChild(elementName);

		var elementLink = document.createElement("span");
		elementLink.setAttribute("class", "elementLink");
		elementLink.setAttribute("recordid", data.id);
		elementLink.innerText = data.name ? data.name : "--";
		elementLink.setAttribute("title", data.name ? data.name : "--");
		elementLink.addEventListener("click", this._linkClicked);
		elementName.appendChild(elementLink);
		
		var elementBody = document.createElement("div");
		elementBody.setAttribute("class", "elementBody");
		element.appendChild(elementBody);

		var elementTable = document.createElement("table");
		elementTable.setAttribute("class", "elementTable");
		elementBody.appendChild(elementTable);

		
		data.columns.forEach(col => {

			var skip = col.isHidden??false;
			if(col.name.endsWith(".stagename")) skip = true;
			if(col.name == "stageid") skip = true;
			if(col.name == this._config.amountAttribute) skip = !this._config.showAmount;
			if(col.name == this._config.nameAttribute) skip = true;
			if(col.name == this._config.picklistAttribute) skip = true;
			if(col.name == this._config.entityAttribute) skip = true;

			if(this._sysColumns.indexOf(col.name) > -1) skip = true;

			if(!skip){
				var tr = document.createElement("tr");
				var td1 = document.createElement("td");
				td1.setAttribute("class", "elementCellTitle");
				td1.innerText = col.displayName;
				td1.setAttribute("title", col.displayName);
				var td2 = document.createElement("td");
				td2.setAttribute("class", "elementCellValue");

				var val = this.createValueElement(col, data.entity);
				td2.appendChild(val);

				tr.appendChild(td1);
				tr.appendChild(td2);
				elementTable.appendChild(tr);
			}

		});

		var saveButton = document.createElement("button");
		saveButton.setAttribute("class", "saveButton");
		saveButton.setAttribute("recordid", data.id);
		saveButton.innerText = this._customText.Save;
		saveButton.addEventListener("click", this._saveElement, true);
		element.appendChild(saveButton);

		element.addEventListener("dragstart", this._dragStart);
		element.addEventListener("dragend", this._dragEnd);

		return element;
	}
	private createValueElement(column:DataSetInterfaces.Column, entity:DataSetInterfaces.EntityRecord):HTMLDivElement{
		var container = document.createElement("div");
		container.setAttribute("recordid", entity.getRecordId());
		try{
			var valueDiv = document.createElement("div");
			valueDiv.setAttribute("class", "elementFormattedValue");
			var value = entity.getValue(column.name);
			var formattedValue = entity.getFormattedValue(column.name);
	
			var filtered = this.attributesMeta.filter(a => a.LogicalName == column.name);
			var dataType = "";
			if(filtered.length > 0) dataType = filtered[0].Type;

			switch(dataType){
				case "Lookup":
				case "Owner":
				case "Customer":
					if(value){
						var d1 = document.createElement("div");
						d1.setAttribute("class", "elementValueLookupPrefix");
						d1.innerText = this.getTextPrefix(formattedValue);
						valueDiv.setAttribute("title", this.getTextPrefix(formattedValue));
						var d2 = document.createElement("div");
						d2.setAttribute("class", "elementValue");
						d2.innerText = formattedValue;
						d2.setAttribute("title", formattedValue);
						valueDiv.appendChild(d1);
						valueDiv.appendChild(d2);
					}else{
						valueDiv.innerText = "--";	
						valueDiv.setAttribute("title", "--");
					}
					break;
				default:
					valueDiv.innerText = ((value || value == "0") ? formattedValue : "--");
					valueDiv.setAttribute("title", ((value || value == "0") ? formattedValue : "--"));
				}

			container.appendChild(valueDiv);
	
			if(this._config.editableAttributes.indexOf(column.name) > -1){
				valueDiv.setAttribute("class", "elementFormattedValue editableElementFormattedValue");

				valueDiv.addEventListener("click", this._valueClicked);
				var editorElement = this.getValueEditor(column, entity);
				container.appendChild(editorElement);
			}
		}catch(exc){
			console.log("createValueElement : " + exc.message);
		}
		return container;
	}
	private getTextPrefix(text:string){
		if(text){
			var t = text.split(" ");
			if(t.length > 1){
				return t[0][0].toUpperCase() + t[1][0].toUpperCase();
			}else if(text.length >1){
				return text[0].toUpperCase() + text[1].toUpperCase();
			}else if(text.length > 0)
				return text[0].toUpperCase();
		}
		return "";
	}
	private getValueEditor(column:DataSetInterfaces.Column, entity:DataSetInterfaces.EntityRecord):HTMLElement{
		try{
			var value = entity.getValue(column.name) ?? "";

			var filtered = this.attributesMeta.filter(a => a.LogicalName == column.name);
			var dataType = "";
			if(filtered.length > 0) dataType = filtered[0].Type;

			if(dataType == "String" || dataType == "Memo"){
				var e1 = document.createElement("input");
				e1.setAttribute("class", "elementValueEditor");
				e1.setAttribute("type", "text");
				e1.value = value.toString();
				e1.setAttribute("currvalue", value.toString());
				e1.setAttribute("attribute", column.name);
				e1.setAttribute("datatype", dataType);
				return e1;
			}
			else if(dataType == "Picklist"  || dataType == "State"  || dataType == "Status"){
				var meta = this.allMetadata.filter(a => a.logicalname == column.name);
				if(meta.length > 0){
					var e3 = document.createElement("select");
					var opts = "";
					meta[0].options.forEach(a => {
						opts += "<option value='"+a.value+"'>"+a.label+"</option>";
					});
					e3.innerHTML = opts;
					e3.setAttribute("class", "elementValueEditor");
					e3.value = value.toString();
					e3.setAttribute("currvalue", value.toString());		
					e3.setAttribute("attribute", column.name);
					e3.setAttribute("datatype", dataType);
					return e3;				
				}
			}
			else if(dataType == "Boolean"){
				var meta = this.allMetadata.filter(a => a.logicalname == column.name);
				if(meta.length > 0){
					var e4 = document.createElement("select");
					var opts = "";
					meta[0].options.forEach(a => {
						opts += "<option value='"+a.value+"'>"+a.label+"</option>";
					});
					e4.innerHTML = opts;
					e4.setAttribute("class", "elementValueEditor");
					e4.value = value.toString();
					e4.setAttribute("currvalue", value.toString());						
					e4.setAttribute("attribute", column.name);
					e4.setAttribute("datatype", dataType);
					return e4;
				}
			}

			else if(dataType == "Integer"){
				var e5 = document.createElement("input");
				e5.setAttribute("class", "elementValueEditor");
				e5.setAttribute("type", "number");
				e5.setAttribute("step", "1");
				e5.value = value.toString();
				e5.setAttribute("currvalue", value.toString());			
				e5.setAttribute("attribute", column.name);
				e5.setAttribute("datatype", column.dataType);
				return e5;
			}
			else if(dataType == "Decimal" || dataType == "Money"){
				var e6 = document.createElement("input");
				e6.setAttribute("class", "elementValueEditor");
				e6.setAttribute("type", "number");
				e6.setAttribute("step", "0.01");
				e6.value = value.toString();
				e6.setAttribute("currvalue", value.toString());			
				e6.setAttribute("attribute", column.name);
				e6.setAttribute("datatype", dataType);
				return e6;
			}
			else if(dataType == "DateTime"){
				if(column.dataType && column.dataType.indexOf("DateAndTime.DateAndTime") > -1){
					var e7 = document.createElement("input");
					e7.setAttribute("class", "elementValueEditor");
					e7.setAttribute("type", "datetime-local");
					e7.setAttribute("step", "1");
					e7.value = value.toString().replace(".000Z", "");
					e7.setAttribute("currvalue", value.toString().replace(".000Z", ""));			
					e7.setAttribute("attribute", column.name);
					e7.setAttribute("datatype", dataType);
					return e7;
				}else{
					var e7 = document.createElement("input");
					e7.setAttribute("class", "elementValueEditor");
					e7.setAttribute("type", "date");
					e7.setAttribute("step", "1");
					e7.value = value.toString().replace(".000Z", "");
					e7.setAttribute("currvalue", value.toString().replace(".000Z", ""));			
					e7.setAttribute("attribute", column.name);
					e7.setAttribute("datatype", dataType);
					return e7;
				}
			}
			
			debugger;
			return document.createElement("div");
		}
		catch(exc){
			debugger;
			console.log("getValueEditor: " + exc.message);
			return document.createElement("div");
		}
	}
	private valueClicked(evt:Event){
		var target = (evt.currentTarget as HTMLDivElement).parentElement as HTMLDivElement;
		var currentRecordId = target.getAttribute("recordid") ?? "";

		if(this.editedRecordId && this.editedRecordId != currentRecordId){
			this.cancelEdition(this.editedRecordId);	
		}

		this.editedRecordId = currentRecordId;

		var children = target.childNodes;
		var formattedElement = (children[0] as HTMLElement);
		var editorElement = (children[1] as HTMLElement);
		formattedElement.setAttribute("style", "display: none");
		editorElement.setAttribute("style", "display: block");
		editorElement.setAttribute("value", editorElement.getAttribute("currvalue")??"");
		document.querySelectorAll(".saveButton[recordid='" + this.editedRecordId + "']").forEach(a => {
			a.setAttribute("style", "display: block");
		});
	}
	private findGetParameter(parameterName:string) {
		var result = null,
			tmp = [];
		location.search
			.substr(1)
			.split("&")
			.forEach(function (item) {
			  tmp = item.split("=");
			  if (tmp[0] === parameterName) result = decodeURIComponent(tmp[1]);
			});
		return result;
	}
	private linkClicked(evt:Event){
		var target = evt.currentTarget as HTMLDivElement;
		var currentRecordId = target.getAttribute("recordid") ?? "";
		var appid = this.findGetParameter("appid");
		var url = "/main.aspx?etn=" + this._entity + "&id=" + currentRecordId + "&pagetype=entityrecord";
		if(appid)
			url += "&appid=" + appid;

		window.open(url, "_blank");
	}
	private moreClicked(evt:Event){
		var target = evt.currentTarget as HTMLDivElement;
		var groupid = target.getAttribute("groupid")??"";
		var records = this.dataList.filter(a => a.groupid == groupid);

		var section = document.querySelector(".sectionBody[groupid='" + groupid + "']");
		var container = document.querySelector(".sectionBody[groupid='" + groupid + "'] .elementsContainer") as HTMLDivElement;
		var counterCell = document.querySelector(".sectionBody[groupid='" + groupid + "'] td[counter='true']") as HTMLDivElement;
		var maxValue = parseInt(section?.getAttribute("max")??"0");
		var max = maxValue + this._config.maxRecords;

		target.remove();

		var createMoreButton = false;
		var count = maxValue;
		for(var i = maxValue; i < records.length; i++){
			if(i < max){
				count ++;
				container.appendChild(records[i].DivElement as HTMLDivElement);
			}else{
				createMoreButton = true;
			}
		}
		section?.setAttribute("max", count.toString());
		counterCell.innerText = this.formatNumber(count) + " / " + this.formatNumber(records.length);

		if(createMoreButton){
			container.appendChild(this.createMoreButton(groupid));
		}

	}
	private saveElement(evt:Event){
		var target = evt.currentTarget as HTMLElement;
		var id = target.getAttribute("recordid");

		var values = document.querySelectorAll(".element[recordid='" + id + "'] .elementValueEditor");
		var obj = {};
		values.forEach(el => {
			this.setEntityValueForUpdate(obj, el);
		});

		this.showLoader(true);
		var self = this;
		this._context.webAPI.updateRecord(this._entity, this.editedRecordId, obj).then(
			function(){
				self._context.parameters.sampleDataSet.refresh();
			},function(errorResponse:any){
				debugger;
				self.showLoader(false);
				self.showError(errorResponse);
			}
		);
	}
	private setEntityValueForUpdate(obj:any, element:Element){
		var type = element.getAttribute("datatype");
		var attribute = element.getAttribute("attribute");
		var input = null;
		var value = null;
		switch(type){
			case "DateTime":
				input = (element as HTMLInputElement);
				value = input.validity.badInput ? null : input.value;
				//@ts-ignore
				obj[attribute] = value ? value : null;
				break;
			case "Money":
			case "Decimal":
				input = (element as HTMLInputElement);
				value = input.validity.badInput ? null : input.value;
				//@ts-ignore
				obj[attribute] = value ? parseFloat(value) : null;
				break;
			case "Integer":
			case "State":
			case "Status":
			case "Picklist":
				input = (element as HTMLInputElement);
				value = input.validity.badInput ? null : input.value;
				//@ts-ignore
				obj[attribute] = value ? parseInt(value) : null;
				break;
			case "Boolean":
				input = (element as HTMLSelectElement);
				value = input.validity.badInput ? null : input.value;
				//@ts-ignore
				obj[attribute] = (value == "1" ? true : (value == "0" ? false : null));
				break;
						
			case "String":
			case "Memo":
				input = (element as HTMLSelectElement);
				value = input.validity.badInput ? null : input.value;
				//@ts-ignore
				obj[attribute] = value ? value : null;;
				break;
			default:
				debugger;
		}
	}

	private formatNumber(number: number){
		return number.toString().replace(/\B(?=(\d{3})+(?!\d))/g, this._context.userSettings.numberFormattingInfo.currencyGroupSeparator);
	}
	private cancelEdition(recordId:string){
		document.querySelectorAll("div[recordid='" + recordId + "'] .elementFormattedValue").forEach(a => {
			a.setAttribute("style", "");
		});
		document.querySelectorAll("div[recordid='" + recordId + "'] .elementValueEditor").forEach(a => {
			a.setAttribute("style", "");
		});
		document.querySelectorAll(".saveButton[recordid='" + recordId + "']").forEach(a => {
			a.setAttribute("style", "");
		});
	}

	private showLoader(visible:boolean){
		if(visible){
			this._loader.style.display = "block";
		}else{
			this._loader.style.display = "none";
		}
	}
	private createLoader(){
		this._loader = document.createElement("div");
		this._loader.setAttribute("class", "loader");

		this._loader.innerHTML = 
		"<table style='width:100%; height:100%; table-layout: fixed;'>\
			<tr>\
				<td style='text-align: center; vertical-align: middle'><div class='loaderImg'></div></td>\
			</tr>\
		</table>";

		this._container.appendChild(this._loader);
	}

	private getColor(index:number, transparency: string): string{
		var idx = index % this._config.colorsList.length;
		return this._config.colorsList[idx] + transparency;
	}

	private showError(error: string){
		var errorContainer = document.createElement("div");
		errorContainer.setAttribute("class", "errorContainer");

		var errorTable = document.createElement("table");
		errorTable.setAttribute("class", "errorTable");
		errorContainer.appendChild(errorTable);

		var errorTr = document.createElement("tr");
		errorTable.appendChild(errorTr);

		var errorImg = document.createElement("td");
		errorImg.setAttribute("class", "errorImg");
		errorTr.appendChild(errorImg);
		
		var errorTxt = document.createElement("td");
		errorTxt.setAttribute("class", "errorTxt");
		errorTr.appendChild(errorTxt);
		
		var errorTitle = document.createElement("div");
		errorTitle.setAttribute("class", "errorTitle");
		errorTitle.innerText = "Ooops !";
		errorTxt.appendChild(errorTitle);

		var errorMessage = document.createElement("div");
		errorMessage.setAttribute("class", "errorMessage");
		errorMessage.innerText = error;
		errorTxt.appendChild(errorMessage);

		this._container.innerHTML = "";
		this._container.appendChild(errorContainer);
	}


	private udateRecord(entity:string, entityId: string, obj: any, callback: Function){
		//@ts-ignore
		var serverUrl = Xrm.Page.context.getClientUrl();
		//@ts-ignore
		var version:string = Xrm.Page.context.getVersion();
		var v = version.split(".");
		var apiVersion = [v[0],v[1]].join(".");

		var request = serverUrl + "/api/data/v" + apiVersion + "/" + entity + "(" + entityId + ")";
	
		var req = new XMLHttpRequest();
		req.open("PATCH", request);
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.onreadystatechange = function() {
			if (this.readyState === 4) {
				req.onreadystatechange = null;
				if (this.status === 204) {
					callback(true);
				} else {
					debugger;
					callback(false, this.statusText);
				}
			}
		};
		req.send(JSON.stringify(obj));
	}
}

enum GroupType{
	Entity=293160000,
	Picklist=293160001,
	BPF=293160002,
}
interface ConfigObject{
	processId: string,
	amountAttribute: string, 
	nameAttribute: string,
	showAmount: boolean,
	currencyId: string,
	editableAttributes: Array<string>,
	colorsList: Array<string>,
	colorfullsections: boolean,
	minWidthSection: number,
	groupType: GroupType
	bpfEntity: string,
	bpfRecordAttribute: string,
	picklistAttribute: string,
	entityLogicalName: string,
	entityIdAttribute: string,
	entityNameAttribute: string,
	entityFetchXml: string,
	entityAttribute: string,
	maxRecords: number,
	ShowEmptyGroup: boolean,
	EntityIcon: string
}
interface DataObject {
	id: string,
	name: string,
	amount: number,
	groupid: string,
	columns: Array<DataSetInterfaces.Column>,
	entity: DataSetInterfaces.EntityRecord,
	DivElement: HTMLDivElement|undefined
}
interface GroupObject{
	order: number,
	id: string,
	name: string,
	DropContainer: HTMLDivElement|undefined
}
interface RecordStageObject{
	stageid: string,
	recordid: string,
	id: string
}
interface EntityMeta{
	LogicalName: string,
	LogicalCollectionName: string
}
interface AttributeMeta{
	LogicalName: string,
	Type: string
}
interface OptionSetMeta{
	logicalname: string,
	type: string,
	options: Array<OptionMeta>
}
interface OptionMeta{
	value: number,
	label: string,
	parentValue: number|undefined
}
interface CustomText{
	Save: string,
	More: string,
	Empty: string
}
interface Currency{
	id: string,
	rate: number,
	symbol: string
}