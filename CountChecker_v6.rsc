/*
TMIP Count Checker Tool:

OBJECTIVE:
	Identify count inconsistency problems in a network

FUNCTIONS:
	count coverage for the network
	capacity-based check
	count propagation
	intersection level checks
	intersection turning movement calculations
	intersection missing AADT calculations

GUI:
	import settings
	view a report from the GUI
	create maps for a report

INPUTS:
	Highway network database
	Fields:
		LinkID
		Length
		Functional Class
		Count AB/BA
		Capacity AB/BA

OUTPUTS:
	Reports Summary						[output_dir]\ReportsSummary.txt
	Count Coverage						[output_dir]\CountCoverage.txt
	Count Propagation					[output_dir]\LinksWithPropagatedCounts.csv
	Capacity-based Checks				[output_dir]\LinkCapacityBasedChecks.csv
	Intersection Flow Conservation		[output_dir]\IntersectionFlowConsCheck.csv
	Intersection Turning Movements		[output_dir]\IntersectionTurnMovements.csv
	Intersection AADT Calculations		[output_dir]\IntersectionCalculatedCount.csv
	Intersection Missing AADT			[output_dir]\IntersectionMissingCount.csv

AUTHOR:
	nagendra.dhakar@rsginc.com
	Updated: 05/19/2016
	Updated: 03/19/2018 (minor bug related to null selection query)
*/

Macro "run"
	manager = CreateObject("CountCheckerGUI.Controller",null)
	output_data=RunDBox("CountChecker", manager)

EndMacro

Class "CountCheckerGUI.Controller"

	init do
		shared info, inputs, outputs, fields
		
		on escape, notfound do
			on escape default
			on notfound default
			return()
		end
		
		mapname = GetMap()
		if mapname<>null then SetMapRedraw(null, "False")
		SetSelectDisplay("False")
	
	endItem
	
endClass

DBox "CountChecker" (controller) title: " TMIP Count Checker"
	
	Toolbox
		Nokeyboard
		Docked: right
	
	init do
		shared inputs, outputs, fields, bool, thresholds, list_fields, index, msg, delim, lst_reports, lst_results, layers, fields_display
		shared set_res_idx, idx
		
		fields={}
		inputs={}
		outputs={}
		reports={}
		bool={}
		thresholds={}
		msg = {}
		index = {}
		delim = "="
		set_res_idx = 0
		idx=0
	
	endItem
	
	button "Close" 43.5, 43, 10, 2 help: "Exit the tool" Cancel do
		mapname = GetMap()
		if mapname<>null then SetMapRedraw(null, "True")
		return()
	endItem
	
	Tab list 1, 0.5, 57, 46 variable: tab_idx
	
	Tab prompt: "Run Tool"
		
		init do
		
			msg.title = "TMIP Count Checker v0"
			msg.date = "May 19, 2016"
			msg.company = "Resource Systems Group, Inc."
			msg.author = "Nagendra S Dhakar and John Gliebe"
			bool.uturn=0
			bool.twowayinter=0
			thresholds.caplow = 0.1  		// percent of capacity - for count vs capacity check
			thresholds.caphigh = 0.9 		// percent of capacity - for count vs capacity check
			thresholds.aadt = 0.9       // ratio of inbound and outbound AADT - for intersection flow conservation
			thresholds.countlow = 0.1
			thresholds.counthigh = 0.9
			thresholds.gap = 0.5			// sum of column gap during intersection turning movement balancing
			list_reports = {"count coverage","capacity-based check", "count propogation", "intersection flow conservation", 
			"intersection turning movements", "calculate missing intersection count", "intersections with missing counts"}
		endItem

		// Select input settings
		button "settingsbrowse" 1.5, 0.5, 15, 1.5 prompt: "Import Settings" help: "Import tool inputs from a text file" do on escape goto endhere
			inputs.settings = ChooseFile({{"Input Settings", "*.txt"}}, "Choose inputs settings", {,{"Input Settings", inputs.path},})
			RunMacro("Populate GUI")
		endhere: endItem
		
		text "Region" 2, 3.25
		edit text "scenname" 15, same, 15 variable: msg.region do 
			endItem

		text "Year" 35, 3.25
		edit Int "year" 45, same, 6 variable: msg.year do
			endItem
		
		// Select project directory
		Frame 1.5, 5, 53, 3.5 prompt: "Select 1: Project Directory"
		button "pathbrowse" 44, 6.5, 8 prompt: "Browse" help: "Select project directory"  do on escape goto endhere
			inputs.path = ChooseDirectory("Choose the Model Directory", {{"Initial Directory", inputs.path}})
		endhere: endItem
		edit text "projdir" 3, 6.5, 39.5 variable: inputs.path do
			endItem
		
		Frame 1.5, 9, 53, 5 prompt: "Select 2: Network Database"		
		// Select network database
		button "networkbrowse" 44, 10.5, 8 prompt: "Browse" help: "Select network database" do on escape goto endhere
			inputs.dbfile = ChooseFile({{"Master Network", "*.dbd"}}, "Choose the masternet layer", {,{"Initial Directory", inputs.path},})
			list_fields = RunMacro("Read Network Fields")
		endhere: endItem
		edit text "network" 3, 10.5, 39.5 variable: inputs.dbfile do
			endItem
		
		text "Selection Query" 3.5, 12.5 help: "To analyze a selected region in the network (Optional)"
		edit text "query" 17.5, same, 25 variable: inputs.query do
			endItem
		text "(optional)" 43.5, same
		
		// Select fields
		Frame 1.5, 14.5, 53, 10 prompt: "Select 3: Database Fields"
		// drop down menus for fields
		Popdown Menu "linkid" 17, 16.5, 10 prompt: "Link ID" list: list_fields variable: index.linkid do on escape goto endhere
			fields.linkid = list_fields[index.linkid]
		endhere: endItem
		Popdown Menu "countab" 42, same, 10 prompt: "Count AB" list: list_fields variable: index.count_AB do on escape goto endhere
			fields.count_AB = list_fields[index.count_AB]
		endhere: endItem		
		Popdown Menu "direction" 17, 18.5, 10 prompt: "Link Direction" list: list_fields variable: index.direction do on escape goto endhere
			fields.direction = list_fields[index.direction]
		endhere: endItem
		Popdown Menu "countba" 42, same, 10 prompt: "Count BA" list: list_fields variable: index.count_BA do on escape goto endhere
			fields.count_BA = list_fields[index.count_BA]
		endhere: endItem	
		Popdown Menu "length" 17, 20.5, 10 prompt: "Length" list: list_fields variable: index.len do on escape goto endhere
			fields.len= list_fields[index.len]
		endhere: endItem
		Popdown Menu "capacityab" 42, same, 10 prompt: "Capacity AB" list: list_fields variable: index.capacity_AB do on escape goto endhere
			fields.capacity_AB = list_fields[index.capacity_AB]
		endhere: endItem
		Popdown Menu "funcclass" 17, 22.5, 10 prompt: "Functional Class" list: list_fields variable: index.funcclass do on escape goto endhere
		fields.funcclass = list_fields[index.funcclass]
		endhere: endItem
		
		Popdown Menu "capacityba" 42, same, 10 prompt: "Capacity BA" list: list_fields variable: index.capacity_BA do on escape goto endhere
			fields.capacity_BA= list_fields[index.capacity_BA]
		endhere: endItem	
		
		// Other settings
		Frame 1.5, 25, 53, 11 prompt: "Select 4: Other Settings"
		checkbox 3, 26.5 prompt: "Allow U-turns?" Disabled variable: bool.uturn
		checkbox 25, same prompt: "Process 2-way intersections?" Disabled variable: bool.twowayinter 
		
		text "Capacity Range Factors:  " 3, 28.5 help: "Factors used to calculate range of capacity to compare with counts"
		
		text "Low" 29, same
		edit Real "thres_caplow" 33.5, same, 6 variable: thresholds.caplow help: "Factor applied to capacity to find lower range of count" do
			endItem		
		
		text "High" 41.5, same
		edit Real "thres_caphigh" 46, same, 6 variable: thresholds.caphigh help: "Factor applied to capacity to find higher range of count" do
			endItem	

		text "Missing Count Range Factors:  " 3, 30.5 help: "Factors used to calculate a range of AADT count on an approach"
		
		text "Low" 29, same
		edit Real "thres_countlow" 33.5, same, 6 variable: thresholds.countlow do
			endItem		
		
		text "High" 41.5, same
		edit Real "thres_counthigh" 46, same, 6 variable: thresholds.counthigh do
			endItem	
			
		text "Intersection Count Ratio Threshold" 3, 32.5 help: "Threhshold on the ratio of an inbound flow and total outbound from other legs"
		edit Real "thres_aadt" 33.5, same, 6 variable: thresholds.aadt do
			endItem
			
		text "Turning Movements Error Tolerance" 3, 34.5 help: "Acceptable gap in turning movement calculations"
		edit Real "thres_gap" 33.5, same, 6 variable: thresholds.gap do
			endItem
		
		Frame 1.5, 36.5, 53, 3.5 prompt: "Select 5: Output Directory"
		
		// Select output directory
		button "outpathbrowse" 44, 38, 8 prompt: "Browse" help: "Select output directory"  do on escape goto endhere
			outputs.dir = ChooseDirectory("Choose the Model Directory", {{"Initial Directory", inputs.path}})			
		endhere: endItem
		edit text "outdir" 3, same, 39.5 variable: outputs.dir do
			endItem

		// Run the tool
		Button "run" 16.5, 41, 22, 2 prompt: "Run Tool" help: "Run count checks" do on escape goto endhere
			if fields.linkid=null or fields.direction=null or fields.funcclass=null  or fields.len=null or fields.count_AB=null 
			or fields.count_BA=null or fields.capacity_AB=null or fields.capacity_BA=null then ShowMessage("Error: one or more fields are not selected.")
			else do
				RunMacro("Run Tool")
			end
		endhere: endItem

	Tab prompt: "Results"		
		Frame 2, 8, 52, 6 prompt: "View Report"
		
		Popdown Menu "list_reports" 6, 11, 28 list: lst_reports variable: set_rptidx do on escape goto endhere
			chosen_report = lst_reports[set_rptidx]
		endhere: endItem		

		Button "view_report" 40, 11, 12, 1 prompt: "View" help: "Open the selected report" do on escape goto endhere
			if len(chosen_report) > 0 then RunMacro("Open Report",{chosen_report})
			else MessageBox("Choose a report",)
		endhere: endItem		
		
		Frame 2, 18, 52, 8 prompt: "Display Report"
		
		Popdown Menu "list_results" 6, 21, 28 list: lst_results variable: set_res_idx do on escape goto endhere
			chosen_result = lst_results[set_res_idx]
		endhere: endItem
		
		text "Field" 40, 19.5
		Popdown Menu "result_fields" 38, 21, 12 list: fields_display[set_res_idx] variable: idx do on escape goto endhere
			chosen_fld = fields_display[set_res_idx][idx]
		endhere: endItem			

		Button "display_result" 23, 24, 12, 1 prompt: "Display" help: "Create a map of results" do on escape goto endhere
			if len(chosen_fld) > 0 then do
				// if a field is chosen
				if ArrayPosition(fields_display[set_res_idx],{chosen_fld},)=0 then do
					// for cases when report is changed but not a field. 
					// this causes chosen field to be from the last run, which may not be in the current chosen report
					// so, make sure that the field is in the current report, if not then prompt to choose a field				
					MessageBox("Choose a field",)
				end
				else RunMacro("Display Result", chosen_result, chosen_fld)
			end
			else if len(chosen_result) >0 then do
				// if report is chosen but not field
				// prompt to choose a field
				MessageBox("Choose a field",)
			end
			else MessageBox("Choose a report",)
		endhere: endItem	
		
/*	
	Tab prompt: "Intersection Balancing"
	
		text "NORTH" 28, 2
		text "R" 22, 3.5
		text "T" 30.5, same
		text "L" 39, same
		
		edit int "northright" 19, 5, 7 variable: north.right do endItem
		edit int "norththrough" 27.5, same, 7 variable: north.through do endItem
		edit int "northleft" 36, same, 7 variable: north.left do endItem

		//text "Hold cross-street counts constant" 30, 1
		checkbox 30, 0.5 prompt: "Hold constant?" variable: bool.holdnorth
		
		text "SOUTH" 28, 16.5
		text "L" 22, 15
		text "T" 30.5, same
		text "R" 39, same
		
		edit int "southleft" 19, 13.5, 7 variable: south.left do endItem
		edit int "souththrough" 27.5, same, 7 variable: south.through do endItem
		edit int "southright" 36, same, 7 variable: south.right do endItem
		
		checkbox 30, 18 prompt: "Hold constant?" variable: bool.holdsouth
		
		text "WEST" 3, 9
		text "L" 9,7 
		text "T" same, 9
		text "R" same, 11
		
		edit int "westleft" 11, 7, 7 variable: west.left do endItem
		edit int "westthrough" same, 9, 7 variable: west.through do endItem
		edit int "westright" same, 11, 7 variable: west.right do endItem

		text "EAST" 54, 9
		text "R" 52.5,7 
		text "T" same, 9
		text "L" same, 11
		
		edit int "eastleft" 44, 7, 7 variable: east.left do endItem
		edit int "eastthrough" same, 9, 7 variable: east.through do endItem
		edit int "eastright" same, 11, 7 variable: east.right do endItem
		
		text "Adj. Intersection" 57.5, 4.5
		frame "adjint" 59, 5.5, 12.5, 7.5 prompt: ""
		
		text "Arrival" 62, 6.5
		edit int "adjintarrive" 62, 8, 7 variable: adjint.arrive
		text "Departure" 62, 10
		edit int "adjintdepart" same, 11.5, 7 variable: adjint.depart
*/		
	
	Tab prompt: "About"
			
		text "Title" 8, 13		
		text "Date" same, 14
		text "Author" same, 15
		text "Company" same, 16
		
		text ":" 16, 13
		text ":" same, 14
		text ":" same, 15
		text ":" same, 16
		
		text 18, 13 variable: msg.title
		text same, 14 variable: msg.date		
		text same, 15 variable: msg.author
		text same, 16 variable: msg.company	

EndDBox

Macro "Run Tool"
	shared layers
	
	RunMacro("Set output settings")
	RunMacro("Write Run Settings")
	RunMacro("Select Region", layers.link)
	RunMacro("Compute Network Stats")
	RunMacro("Get Link Data")
	RunMacro("Check count and capacity")
	RunMacro("Propagate Counts")
	RunMacro("Intersection Flow Conservation")
	RunMacro("Write Summary of Outputs")
	RunMacro("Count Coverage")

EndMacro

Macro "Write Run Settings"
	shared inputs, outputs, bool, thresholds, fields, msg, delim
	
	ptr = OpenFile(outputs.settings,"w")
	
	// write settings to a file
	WriteLine(ptr, JoinStrings({"region",msg.region},delim))
	WriteLine(ptr, JoinStrings({"year",msg.year},"="))
	WriteLine(ptr, JoinStrings({"projdir",inputs.path},delim))
	WriteLine(ptr, JoinStrings({"networkdb",inputs.dbfile},delim))
	WriteLine(ptr, JoinStrings({"query",inputs.query},delim))
	WriteLine(ptr, JoinStrings({"field_linkid",fields.linkid},delim))
	WriteLine(ptr, JoinStrings({"field_dir",fields.direction},delim))
	WriteLine(ptr, JoinStrings({"field_func",fields.funcclass},delim))
	WriteLine(ptr, JoinStrings({"field_length",fields.len},delim))
	WriteLine(ptr, JoinStrings({"field_countab",fields.count_AB},delim))
	WriteLine(ptr, JoinStrings({"field_countba",fields.count_BA},delim))
	WriteLine(ptr, JoinStrings({"field_capab",fields.capacity_AB},delim))
	WriteLine(ptr, JoinStrings({"field_capba",fields.capacity_BA},delim))
	WriteLine(ptr, JoinStrings({"uturn",i2s(bool.uturn)},delim))
	WriteLine(ptr, JoinStrings({"twowayinter",i2s(bool.twowayinter)},delim))
	WriteLine(ptr, JoinStrings({"thres_caplow",r2s(thresholds.caplow)},delim))
	WriteLine(ptr, JoinStrings({"thres_caphigh",r2s(thresholds.caphigh)},delim))
	WriteLine(ptr, JoinStrings({"thres_countlow",r2s(thresholds.countlow)},delim))
	WriteLine(ptr, JoinStrings({"thres_counthigh",r2s(thresholds.counthigh)},delim))
	WriteLine(ptr, JoinStrings({"thres_aadt",r2s(thresholds.aadt)},delim))
	WriteLine(ptr, JoinStrings({"thres_gap",r2s(thresholds.gap)},delim))
	WriteLine(ptr, JoinStrings({"outdir",outputs.dir},delim))
	
	CloseFile(ptr)

EndMacro

Macro "Populate GUI"
	shared inputs, outputs, bool, thresholds, fields, list_fields, msg, index, delim
	
	ptr = OpenFile(inputs.settings, "r")
	
	settings = {}
	while not FileAtEOF(ptr) do
		lines=ReadLine(ptr)
		subs = ParseString(lines,delim,)
		
		if subs[1] = "query" then subs[2] = SubString(lines,7,StringLength(lines)) // query may contain '=' thus, treat this differently. get all text after first '=' sign
		
		settings.(subs[1]) = subs[2]
		
	end
	
	CloseFile(ptr)
	
	// set global variables
	msg.region = settings.region
	msg.year = s2i(settings.year)
	inputs.path = settings.projdir
	inputs.dbfile = settings.networkdb
	inputs.query = settings.query
	fields.linkid = settings.field_linkid
	fields.direction = settings.field_dir
	fields.funcclass = settings.field_func
	fields.len = settings.field_length
	fields.count_AB = settings.field_countab
	fields.count_BA = settings.field_countba
	fields.capacity_AB = settings.field_capab
	fields.capacity_BA = settings.field_capba
	bool.uturn = s2i(settings.uturn)
	bool.twowayinter = s2i(settings.twowayinter)
	thresholds.caplow = s2r(settings.thres_caplow)
	thresholds.caphigh = s2r(settings.thres_caphigh)
	thresholds.countlow = s2r(settings.thres_countlow)
	thresholds.counthigh = s2r(settings.thres_counthigh)	
	thresholds.aadt = s2r(settings.thres_aadt)
	thresholds.gap = s2r(settings.thres_gap)
	outputs.dir = settings.outdir
	
	// get network fields
	list_fields = RunMacro("Read Network Fields")
	
	// find field index
	index.linkid = ArrayPosition(list_fields,{fields.linkid},)
	index.direction = ArrayPosition(list_fields,{fields.direction},)
	index.funcclass = ArrayPosition(list_fields,{fields.funcclass},)
	index.len = ArrayPosition(list_fields,{fields.len},)
	index.count_AB = ArrayPosition(list_fields,{fields.count_AB},)
	index.count_BA = ArrayPosition(list_fields,{fields.count_BA},)
	index.capacity_AB = ArrayPosition(list_fields,{fields.capacity_AB},)
	index.capacity_BA = ArrayPosition(list_fields,{fields.capacity_BA},)
	
EndMacro

Macro "Set output settings"
	shared outputs, msg, ptr, reports, lst_reports, lst_results, labels, fields_display
	shared labels
	
	// write file handlers
	ptr = {}
	labels = {}
	
	// tool inputs
	outputs.settings = outputs.dir + JoinStrings({"\\inputs_",msg.region,i2s(msg.year),".txt"},"")
	
	//ODME related outputs
	outputs.hwynet = outputs.dir + "\\hwy.net"
	outputs.demandmat = outputs.dir + "\\Demand.mtx"
	outputs.flowtable = outputs.dir + "\\linkflow.bin"
	
	// report files
	outputs.log = outputs.dir + "\\tcadlog.txt"
	outputs.capcheck = outputs.dir + "\\LinkCapacityBasedChecks.csv"
	outputs.propagatedcounts = outputs.dir + "\\LinksWithPropagatedCounts.csv"
	outputs.intercheck = outputs.dir + "\\IntersectionFlowConsCheck.csv"
	outputs.interturns = outputs.dir + "\\IntersectionTurnMovements.csv"
	outputs.intercalc = outputs.dir + "\\IntersectionCalculatedCount.csv"
	outputs.intermissing = outputs.dir + "\\IntersectionMissingCount.csv"
	outputs.countcoverage = outputs.dir + "\\CountCoverage.txt"
	outputs.metafile = outputs.dir + "\\ReportsSummary.txt"
	
	// list of reports
	reports.propagatedcounts = "Count Propagation"
	reports.intercalc = "Intersection AADT Calculations"
	reports.intercheck = "Intersection Flow Conservation"
	reports.interturns = "Intersection Turning Movements"
	reports.intermissing = "Intersection Missing AADT"
	reports.capcheck = "Capacity-based Checks"
	reports.countcoverage = "Count Coverage"
	reports.metafile = "Summary of Reports"
	
	// to populate in GUI
	lst_reports = {reports.metafile, reports.countcoverage,reports.capcheck,reports.propagatedcounts,
					reports.intercheck,reports.interturns,reports.intercalc,reports.intermissing} 
	
	// list of results to populate in GUI
	lst_results = {reports.propagatedcounts,
				   reports.capcheck,
				   reports.intercheck,
				   reports.interturns,
				   reports.intercalc,
				   reports.intermissing, 
				   "ODME Shortest Paths"} 
	
	// correspnding fields to populate in GUI	
	fields_display = {{"msg"},
					  {"msg_ab","msg_ba"},
					  {"msg_check1","msg_check2","msg_check3"},
					  {"msg"},
					  {"msg"},
					  {"interid"},
					  {"Tot_Flow"}}

	labels.interturns = {}
	labels.interturns.text={"0,Completed",
							"1,Flows from one or more arms are too high to calculate turning movements",
							"2,Turning movements cannot be calculated. Please check input flows"}
	labels.interturns.len = 3
	labels.interturns.index = 12

	labels.intercheck = {}
	labels.intercheck.text={"0,Passed intersection check",
							"1,Total flow entering the junction is not equal to the total flow exiting the junction",
							"2,Inbound flow is not less than the sum of outbound flows from other legs",
							"3,Ratio of inbound flows and sum of outbound flows from other legs is too high"}
	labels.intercheck.len = 4

	labels.intercalc = {}
	labels.intercalc.text={"1,Calculated a missing inbound count",
						   "2,Cannot calculate a missing inbound count: total outbound is less than total inbound",
						   "3,Calculated a missing outbound count",
						   "4,Cannot calculate a missing outbound count: total inbound is lower than total outbound",
						   "5,Calculated missing approach counts (both inbound and outbound)",
						   "6,Cannot calculate: inbound/outbound flow is not available"}
	labels.intercalc.len = 6

	labels.capcheck = {}
	labels.capcheck.text={"0,No count",
						   "1,Count is reasonable",
						   "2,Count is low",
						   "3,Count is high",
						   "4,Count/capacity is not available",
						   "5,Count is on unexpected direction"}
	labels.capcheck.len = 6	

	labels.propagatedcounts = {}
	labels.propagatedcounts.text={"0,No count",
								  "1,Existing coded count",
								  "2,Propagated count",
								  "3,Conflicting counts"}
	labels.propagatedcounts.len = 4		
	
EndMacro

/*
FUNCCLASS (TDOT - HPMS):
1: Rural Principal Arterial – Interstate
2: Rural Principal Arterial – Other
6: Rural Minor Arterial
7: Rural Major Collector
8: Rural Minor Collector
9: Rural Local   
11: Urban Principal Arterial – Interstate
12: Urban Principal Arterial- Other Freeways & Expressways
14: Urban Principal Arterial – Other
16: Urban Minor Arterial
17: Urban Collector
19: Urban Local
91: 1-Lane Roundabout
92: 2-Lane Roundabout
99: Centroid Connector

NOTES:
1. following should hold to be a valid count:
	a. count < capacity
	b. count < threshold.high*capacity
	c. count > threhshold.low*capacity
*/

Macro "Read Network Fields"
	shared inputs, outputs
	shared fields, thresholds
	shared layers, db_layers
	shared links_array, output
	
	// node and link layers
	layers = {}
	{layers.node, layers.link} = RunMacro("TCB Add DB Layers", inputs.dbfile)
	db_layers.node = inputs.dbfile + "|" + layers.node
	db_layers.link = inputs.dbfile + "|" + layers.link
	
	SetView(layers.link)
	networkfields = GetFields(layers.link,"All")
	
	return(networkfields[1])
	
EndMacro

Macro "Select Region" (lyr)
	shared outputs
	shared fields
	shared region
	
	SetLogFileName(outputs.log)
	AppendToLogFile(0,"select region")
	
	on Error do goto Error end	
	
	lyrs = GetLayerNames()
	// set current layer
	SetLayer(lyr)
	region = {}
	region.nodeview = "NodeSet"
	region.view = "LinkSet"
	query2 = " and " + fields.funcclass + "<>99"
	region.numlinks = RunMacro("Make Selection",{query2, region.view})
	
	AppendToLogFile(0,"finished selecting region!")
	
	ok=1
	return(ok)
	
	Error:
	ShowMessage("Error!")
	ok=0
	return(ok)
	
EndMacro

Macro "Compute Network Stats"
	shared outputs
	shared fields
	shared region, link_stats
	
	vw=GetView()
	region.viewset = vw + "|" + region.view
	lanemiles_array = GetDataVector(region.viewset,"Length",)
	region.lanemiles = VectorStatistic(lanemiles_array, "sum",)	
	
	// Compute Statistics by facility type
	temp_output = outputs.dir + "\\temp_region.csv"
	region.summary = AggregateTable("Region Summary", region.view,"CSV", temp_output,fields.funcclass,
							{{fields.linkid,"count"},
							 {fields.len,"sum"}},null)	
	
	field_names = {fields.funcclass,"N "+fields.linkid,fields.len}
	
	rec = GetFirstRecord(region.summary+"|", null)
	
	link_stats = {}
	while rec <> null do
		
		func_class = i2s(region.summary.(field_names[1]))
		linkcount = region.summary.(field_names[2])
		lanemiles = region.summary.(field_names[3])
		
		if func_class=null then func_class="NA"
		link_stats.(func_class) = {linkcount,lanemiles}
		
		rec = GetNextRecord(region.summary+"|", null, null)
		
	end
	
	CloseView(region.summary)
	DeleteFile(temp_output)
	DeleteFile(outputs.dir + "\\temp_region.CSX")
	DeleteFile(outputs.dir + "\\temp_region.DCC")

EndMacro

Macro "Select Interchanges" (lyr)

	shared region
	shared interchange_array
	
	// select interchanges
	SetLayer(lyr)
	
	//select nodes that are connected to the links in the selected region
	interchanges = SelectByLinks(region.nodeview, "Several", region.view)
	
	// interchange ids
	interchange_array=GetDataVector(region.nodeview,"ID",)

EndMacro

Macro "Get Link Data"
	shared fields, region
	shared links_array

	// save links into an array
	links_array = GetDataVectors(region.view,{fields.linkid,fields.direction,fields.funcclass,fields.count_AB,fields.count_BA,fields.capacity_AB,fields.capacity_BA},)
	
EndMacro

/*
CAPACITY-BASED CHECKS:

Output:
{"linkid", "direction", "funcclass", "count_AB", "count_BA", "capacity_AB", "capacity_BA", "msg_ab", "msg_ba"}

Message Code	Description
0				No count
1				Count is reasonable
2				Count is low
3				Count is high
4				Count/capacity is not available
5				Count is on unexpected direction
*/
Macro "Check count and capacity"
	shared inputs, outputs, fields, reports
	shared links_array, link_index, ptr
	
	on error, notfound do
		DestroyProgressBar()
		return()
	end
	
	CreateProgressBar("Check counts with capacity","True")
	
	AppendToLogFile(0,"start checking counts and capacity")
	
	// create an ouput file and write header
	
	ptr.capcheck = RunMacro("Open File and Write Header", {reports.capcheck})
	
	link_index={}
	link_count = 0
	// loop through network links
	for i=1 to links_array[1].Length do
		perc = RealToInt(100*i/links_array[1].Length)
		UpdateProgressBar("Checking counts with capacity", perc)
		
		// get link attributes
		link_id = links_array[1][i]
		dir	= links_array[2][i]
		func_class = links_array[3][i]
		aadt_ab = links_array[4][i]
		aadt_ba = links_array[5][i]
		cap_ab = links_array[6][i]
		cap_ba = links_array[7][i]
		
		// store linkid index. also set 0 as processing index (used in count smoothing)
		link_index.(String(link_id))={i,0}
		
		// initialize to no counts - message=0
		message_ab="0"
		message_ba="0"
		ratio_ab=0
		ratio_ba=0
		// link represents both directions: dir=0 (both), dir=1(AB), dir=-1(BA)
		if dir = 0 then do
			// for AB direction, both count and capacity is available 
			if aadt_ab<>null and cap_ab<>null then do
				// for BA direction, both count and capacity is available
				if aadt_ba<>null and cap_ba<>null then do
					// check with low and max thresholds
					{ratio_ab, message_ab} = RunMacro("Check threshold",{aadt_ab,cap_ab,"AB"})
					{ratio_ba, message_ba} = RunMacro("Check threshold",{aadt_ba,cap_ba,"BA"})
				end
				// BA direction has data missing - no count/capacity
				else message_ba = "4"
			end
			// no AB direction data, but BA direction data is available
			else if aadt_ba<>null and cap_ba<>null then do
				message_ab = "4" // no count/capacity in AB direction
				{ratio_ba, message_ba} = RunMacro("Check threshold",{aadt_ba,cap_ba,"BA"})
			end
		end
		// link represents either AB or BA
		else do
			// AB direction
			if (dir=1 and aadt_ab>0) then do
				{ratio_ab, message_ab} = RunMacro("Check threshold",{aadt_ab,cap_ab,"AB"})
				cap_ba = 0
			end
			// BA direction
			else if (dir=-1 and aadt_ba>0) then do 
				{ratio_ba, message_ba}= RunMacro("Check threshold",{aadt_ba,cap_ba,"BA"})
				cap_ab=0
			end
			// wrong direction field - count is on unexpected direction
			else error_ab = "5"
			
		end
		
		if message_ab <> "0" or message_ba <> "0" then do
			link_count = link_count + 1
			outstr = JoinStrings({i2s(link_id),i2s(dir),i2s(func_class),r2s(aadt_ab),r2s(aadt_ba),r2s(cap_ab),r2s(cap_ba),r2s(ratio_ab),r2s(ratio_ba),message_ab,message_ba},",")
			RunMacro("Write File", {reports.capcheck, outstr, link_count, null})
		end
		
	end
	
	// close output file
	CloseFile(ptr.capcheck)
	DestroyProgressBar()
	
	AppendToLogFile(0,"finished checking counts and capacity!")
	
	ok=1
	return(ok)
	
EndMacro

Macro "Check threshold" (args)
	shared thresholds
	
	count = args[1]
	capacity = args[2]
	direction = args[3]
	
	ratio = count/capacity
	msg="1" // default to reasonable
	// high threshold check
	if ratio > thresholds.caphigh then msg="3" // count is high
	// low threshold check
	else if ratio < thresholds.caplow then msg="2" // count is low
	
	return({ratio, msg})

EndMacro
/*
PROPAGATED COUNTS:

Output:
{"linkid", "direction", "countAB", "countBA", "message"}

Code   Description
0      No counts
1      Existing count
2 	   Propagated count
3      Count mismatch
*/
Macro "Propagate Counts"
	shared inputs, outputs, fields
	shared layers
	shared links_array, link_index, ptr, reports
	
	//on error goto quit
	
	AppendToLogFile(0,"start propagating counts")
	
	CreateProgressBar("Propagting counts","True")
	
	// create an ouput file and write header
	ptr.propagatedcounts = RunMacro("Open File and Write Header",{reports.propagatedcounts})
	link_count = 0
	for i=1 to links_array[1].length do
	
		perc = RealToInt(100*i/links_array[1].Length)
		UpdateProgressBar("Smoothing counts", perc)
		
		//index=i
		
		// check if link is already processed
		link_id = links_array[1][i]
		array = link_index.(String(link_id))
		isProcessed = array[2]  // if link processed already; 1-yes, 0-no
		
		if isProcessed=0 then do

			trav_links={}
			message=" "
		
			// set link layer as the current layer
			setlayer(layers.link)
			
			// get to and from nodes - starting nodes
			nodes_start = GetEndPoints(link_id)
			
			counts={} // store counts if found
			
			// traverse network in both directions
			for n=1 to nodes_start.length do
				// node
				node=nodes_start[n]
							
				start=True 		// flag for start from the root link
				index=i
				// link attributes
				link_id = links_array[1][index]
				dir	= links_array[2][index]
				aadt_ab = links_array[4][index]
				aadt_ba = links_array[5][index]
				
				// check availability of counts on the link
				if (dir=0 and aadt_ab=null and aadt_ba=null) or (dir=1 and aadt_ab=null) or (dir=-1 and aadt_ba=null) then do 
					isCount=False
				end
				else do 
					isCount=True
					link_index.(String(link_id))={index,1}
					counts.link ={aadt_ab, aadt_ba}
					break
				end				
				
				// traverse network in the direction
				While (!isCount and index<>null) do
					
					if (!start) then do
						// next link attributes
						link_id = links_array[1][index]
						dir	= links_array[2][index]
						aadt_ab = links_array[4][index]
						aadt_ba = links_array[5][index]					
						
						// set link layer as the current layer
						setlayer(layers.link)
						
						// get to and from nodes
						nodes = GetEndPoints(link_id)
						
						FromNode = nodes[1]
						ToNode = nodes[2]
						
						// next node
						if ToNode=node then node=FromNode
						else node=ToNode
						
						// check availability of counts
						if (dir=0 and aadt_ab=null and aadt_ba=null) or (dir=1 and aadt_ab=null) or (dir=-1 and aadt_ba=null) then isCount=False
						else do 
							isCount=True
							
							// save counts by direction
							if n=1 then counts.dir1 ={aadt_ab, aadt_ba}
							else counts.dir2 = {aadt_ab,aadt_ba}
							
							break
						end						
						
					end
					
					// set false
					start=0
					
					// save this link to the array of traversed links
					trav_links.(String(link_id))=index
					
					// and assign 1 as an indicator that the link is processed
					link_index.(String(link_id))={index,1}					
					
					// get set of links on the node
					SetLayer(layers.node)
					link_set = GetNodeLinks(node)
					
					// number of links (approaches)
					nlinks = link_set.length
					
					// check if link set has centroid connector - don't count connectors
					for j=1 to nlinks do
						id = link_set[j]
						array = link_index.(String(id))
						
						// centroid connector
						if array=null then do
							j=j-1
							nlinks=nlinks-1
							pos = ArrayPosition(link_set,{id},)
							link_set = ExcludeArrayElements(link_set,pos,1)
							//link_set = ArrayExclude(link_set,{link_id})
						end
					end
					
					// recalculate number of links
					nlinks = link_set.length			
					
					if (nlinks=1 or nlinks>2) then break
					
					// find next link at the node - the link that is not the current link
					if link_set[1] = link_id then nextlink = link_set[2]
					else nextlink = link_set[1]
					
					// find index of next link
					array = link_index.(String(nextlink))
					
					// if an empty array then set index to null
					if array<> null then index = array[1]
					else index=null
					
				end
			end
			
			if trav_links[1] <> null then do 
				// links are traversed

				if counts.dir1 <> null then do
					message = "2" // count assigned
					aadt_ab=counts.dir1[1]
					aadt_ba=counts.dir1[2]
					if counts.dir2 <> null then do
						if counts.dir1[1]<>counts.dir2[1] and counts.dir1[2] <> counts.dir2[2] then do 
							message = "3" // count mismatch
						end
					end
				end
				else if counts.dir2 <> null then do
					message = "2" // count assigned
					aadt_ab=counts.dir2[1]
					aadt_ba=counts.dir2[2]
				end
				else do
					message = "0" // no count assigned
					// no count
					aadt_ab=null
					aadt_ba=null
				end

				for j=1 to trav_links.length do
					
					// index of the traversed link
					index = trav_links[j][2]
					
					// assign counts
					links_array[4][index] = aadt_ab
					links_array[5][index] = aadt_ba

					// write to an output
					link_count = link_count + 1
					outstr = JoinStrings({i2s(links_array[1][index]),i2s(links_array[2][index]),r2s(links_array[4][index]),r2s(links_array[5][index]),message},",")
					RunMacro("Write File", {reports.propagatedcounts, outstr, link_count, null})
					//writeline(ptr.propagatedcounts,JoinStrings({i2s(links_array[1][index]),i2s(links_array[2][index]),r2s(links_array[4][index]),r2s(links_array[5][index]),message},","))	
					
				end
			end
				
			else if (isCount) then do
				// start link has counts
				message = "1"

				// write to an output
				link_count = link_count + 1
				outstr = JoinStrings({i2s(links_array[1][i]),i2s(links_array[2][i]),r2s(links_array[4][i]),r2s(links_array[5][i]),message},",")
				RunMacro("Write File", {reports.propagatedcounts, outstr, link_count, null})
				//writeline(ptr.propagatedcounts,JoinStrings({i2s(links_array[1][i]),i2s(links_array[2][i]),r2s(links_array[4][i]),r2s(links_array[5][i]),message},","))	
				
			end			
			
		end
	
	end
	
	// close the output file
	CloseFile(ptr.propagatedcounts)
	DestroyProgressBar()
	
	AppendToLogFile(0,"finished propagating counts!")
/*	
	quit:
		DestroyProgressBar()
		return(0)
*/		

EndMacro

Macro "Intersection Flow Conservation" (args)
	shared layers
	shared fields, outputs, thresholds
	shared links_array, link_index, reports
	shared interchange_array, turn_movements, link_set, intertype_array, ptr
	
	//on error, notfound goto quit
	
	AppendToLogFile(0,"start intersection flow conservation")
	
	CreateProgressBar("Intersection flow conservation","True")
	
	// Select interchanges	
	RunMacro("Select Interchanges", layers.node)
	
	// open output files and write labels
	ptr.intercheck = RunMacro("Open File and Write Header", {reports.intercheck})
	ptr.interturns = RunMacro("Open File and Write Header", {reports.interturns})
	ptr.intercalc = RunMacro("Open File and Write Header", {reports.intercalc})
	ptr.intermissing = RunMacro("Open File and Write Header", {reports.intermissing})
	
	{inter_count, line_count, inter_missing, inter_calc} = {0,0,0,0}
	
	// create an array to save nlinks, and count presence (all or missing one approach)
	dim intertype_array[interchange_array.length,2]

	// go through each interchange
	for inter = 1 to interchange_array.length do
	
		perc = RealToInt(100*inter/interchange_array.Length)
		UpdateProgressBar("Checking intersections", perc)	
				
		inter_id = interchange_array[inter]

		if inter_id = 110663 then do
			test=99
		end
		
		// find connected links
		SetLayer(layers.node)
		link_set = GetNodeLinks(inter_id)
		
		// initial number of links (approaches)
		nlinks = link_set.length
		
		// check if link set has centroid connector - don't count connectors
		for i=1 to nlinks do
			link_id = link_set[i]
			array = link_index.(String(link_id))
			
			// centroid connector
			if array=null then do
				i=i-1
				nlinks=nlinks-1
				pos = ArrayPosition(link_set,{link_id},)
				link_set = ExcludeArrayElements(link_set,pos,1)
				//link_set = ArrayExclude(link_set,{link_id})
			end
		end
		
		// recalculate number of links
		nlinks = link_set.length
		
		// set type of intersection: 2-way, 3-way, 4-way, 5-way, ...
		intertype_array[inter][1] = nlinks
		intertype_array[inter][2] = 0   //set default to all counts available
		
		// process 2-way intersections only if switch is set to true
		if (nlinks>2 or bool.twowayinter) then do

			// inbound and outbound links
			// inbound and outbound AADTs
			MissingAADT = False
			dim link_inbound[nlinks], link_outbound[nlinks]
			dim aadt_inbound[nlinks], aadt_outbound[nlinks]
			dim missing_inbound[nlinks], missing_outbound[nlinks] // values: 1-missing, 0-available
			dim heading_inbound[nlinks], heading_outbound[nlinks] // values: from 0 (due north) to 360 (measured clockwise)
			dim turn_movements[nlinks,nlinks] // inbound * outbound
			
			for i=1 to nlinks do
				
				link_id = link_set[i]
				
				array = link_index.(String(link_id))
				
				// default values
				message=""
				{aadt_inbound[i],aadt_outbound[i]} = {0,0}  // set default to 0
				{missing_inbound[i], missing_outbound[i]}={1,1} // set default to missing				
				
				// set default turn movements to 0 - options={ib_idx, ob_idx, initial value}
				RunMacro("Set initial turn movements",{i,0,0}) 
				RunMacro("Set initial turn movements",{0,i,0})
				
				// process only if array is not null
				if array <> null then do
					index = array[1]
					link_dir = links_array[2][index]
					aadt_ab = links_array[4][index]
					aadt_ba = links_array[5][index]
					
					// initialize variables with default values
					InterIsFromNode=False
					
					// get from node
					SetLayer(layers.link)
					nodes = GetEndPoints(link_id)
					from_node = nodes[1]
					
					// check if interchange is from node of the link
					if inter_id=from_node then InterIsFromNode=True
					
					if link_dir=0 then do
						// both direction
						{link_inbound[i], link_outbound[i]}={1,1}
						// heading of the link
						link_heading_ab = Heading(link_id,{{"Direction","AB"}})
						link_heading_ba = Heading(link_id,{{"Direction","BA"}})
						
						if aadt_ba <> null and aadt_ab <> null then do
							{missing_inbound[i], missing_outbound[i]}={0,0}
							RunMacro("Set initial turn movements",{i,0,1}) // inbound turn movements
							RunMacro("Set initial turn movements",{0,i,1}) // outbound turn movements
							if (InterIsFromNode) then do 
								{aadt_inbound[i],aadt_outbound[i]} = {aadt_ba,aadt_ab}
								{heading_inbound[i], heading_outbound[i]} = {link_heading_ba,link_heading_ab}
							end
							else do 
								{aadt_inbound[i],aadt_outbound[i]} = {aadt_ab,aadt_ba}
								{heading_inbound[i], heading_outbound[i]} = {link_heading_ab,link_heading_ba}
							end
						end
						else MissingAADT=True
					end
					else if link_dir=1 then do
						// AB direction
						link_heading_ab = Heading(link_id,{{"Direction","AB"}})
						if aadt_ab <> null then do
							{missing_inbound[i], missing_outbound[i]}={0,0}
							if (InterIsFromNode) then do
								{link_inbound[i], link_outbound[i]}={0,1}
								{aadt_inbound[i],aadt_outbound[i]} = {0,aadt_ab}
								{heading_inbound[i], heading_outbound[i]} = {null,link_heading_ab}
								RunMacro("Set initial turn movements",{0,i,1})
							end
							else do 
								{link_inbound[i], link_outbound[i]}={1,0}
								{aadt_inbound[i],aadt_outbound[i]} = {aadt_ab,0}
								{heading_inbound[i], heading_outbound[i]} = {link_heading_ab,null}
								RunMacro("Set initial turn movements",{i,0,1})
							end
						end
						else do
							if (InterIsFromNode) then do
								{link_inbound[i], link_outbound[i]}={0,1}
								{missing_inbound[i], missing_outbound[i]}={0,1}
							end
							else do
								{link_inbound[i], link_outbound[i]}={1,0}
								{missing_inbound[i], missing_outbound[i]}={1,0}
							end
							MissingAADT=True
						end
					end
					else do
						// BA direction
						link_heading_ba = Heading(link_id,{{"Direction","BA"}})
						
						if aadt_ba <> null then do
							{missing_inbound[i], missing_outbound[i]}={0,0}
							if (InterIsFromNode) then do 
								{link_inbound[i], link_outbound[i]}={1,0}
								{aadt_inbound[i],aadt_outbound[i]} = {aadt_ba,0}
								{heading_inbound[i], heading_outbound[i]} = {link_heading_ba,null}
								RunMacro("Set initial turn movements",{i,0,1})
							end
							else do 
								{link_inbound[i], link_outbound[i]}={0,1}
								{aadt_inbound[i],aadt_outbound[i]} = {0,aadt_ba}
								{heading_inbound[i], heading_outbound[i]} = {null,link_heading_ba}
								RunMacro("Set initial turn movements",{0,i,1})
							end
						end
						else do
							if (InterIsFromNode) then do
								{link_inbound[i], link_outbound[i]}={1,0}
								{missing_inbound[i], missing_outbound[i]}={1,0}
							end
							else do
								{link_inbound[i], link_outbound[i]}={0,1}
								{missing_inbound[i], missing_outbound[i]}={0,1}
							end
							MissingAADT=True
						end			
					end
				end
				
			end
			
			// total count by inbound and outbound
			aadt_inbound_total = Sum(aadt_inbound)
			aadt_outbound_total = Sum(aadt_outbound)			
			
			if (!MissingAADT) then do
				// counts on all approaches
				intertype_array[inter][2] = 0
				inter_count = inter_count + 1
				check2_arr = RunMacro("Intersection Checks",{link_set, aadt_inbound, aadt_outbound, inter_id, nlinks, inter_count})
				
				line_count = line_count + 1
				line_count = RunMacro("Calculate Turn Movements",{link_set, heading_inbound, heading_outbound, aadt_inbound, aadt_outbound, turn_movements, check2_arr, inter_id, nlinks, line_count})
			end
			else do
				// intersection with missing data						
				// total inbound and outbound AADT
				aadt_missing_inbound = sum(missing_inbound)
				aadt_missing_outbound = sum(missing_outbound)
				
				// number of missing counts
				total_missing = aadt_missing_inbound + aadt_missing_outbound
				
				// provide guidance on AADT value only if at-most two directions are missing
				if aadt_missing_inbound <= 1 and aadt_missing_outbound <= 1 then do
					// missing one approach counts
					intertype_array[inter][2] = 1
					inter_calc = inter_calc + 1
					RunMacro("Intersection AADT Calculations", {link_set, missing_inbound, missing_outbound, aadt_inbound, aadt_outbound, inter_id, nlinks, inter_calc})			
				end
				else do
					// missing more than one approach counts
					intertype_array[inter][2] = 2
					inter_missing = inter_missing + 1
					RunMacro("Intersection Missing More",{link_set, missing_inbound, missing_outbound, inter_id, nlinks})			
				end
				
			end
			
		end
	end
	
	// close output file
	CloseFile(ptr.interturns)
	CloseFile(ptr.intercheck)
	CloseFile(ptr.intercalc)
	CloseFile(ptr.intermissing)
	
	DestroyProgressBar()
	
	AppendToLogFile(0,"finished intersection flow conservation!")
	quit:
	return(0)
	DestroyProgressBar()
	
EndMacro

Macro "Intersection Checks" (Args)

	shared ptr, reports, bool, thresholds
	
	linkset = Args[1]
	aadt_ib =Args[2]
	aadt_ob = Args[3]
	interid = Args[4]
	numlinks = Args[5]
	numline = Args[6]
	
	dim check_arr[linkset.length]
	
	//Passed intersection checks - default
	message1="0"
	message2="0"
	message3="0"

	// total count by inbound and outbound
	aadt_ib_total = Sum(aadt_ib)
	aadt_ob_total = Sum(aadt_ob)			
	
	// Check 1. intersection level check
	if aadt_ib_total<>aadt_ob_total then message1="1" //Total flow entering the junction is not equal to the total flow exiting the junction
	
	outstr2="--"
	outstr3="--"
	// Check 2: intersection approach level-check
	for i=1 to linkset.length do
		check_arr[i] = 0
		aadt_leg = aadt_ib[i]
		linkid = linkset[i]
		if aadt_leg>0 then do
			if (!bool.uturn) then do
				// 2.1: an inbound AADT should be less than the sum of outbound AADTs
				if aadt_leg > (aadt_ob_total - aadt_ob[i]) then do 
					message2 = "2" //Inbound flow is not less than the sum of outbound flows from other legs
					check_arr[i] = 1
					if outstr2="--" then outstr2 = i2s(linkid)
					else outstr2 = JoinStrings({outstr2,i2s(linkid)},";")
				end
				
				// 2.2: ratio of an inbound flow and sum of outbound AADTs should be less than 0.9
				ratio = aadt_leg/(aadt_ob_total - aadt_ob[i])
				if ratio > thresholds.aadt then do 
					message3 = "3" //Ratio of inbound flows and sum of outbound flows from other legs is too high
					if outstr3="--" then outstr3 = i2s(linkid)
					else outstr3 = JoinStrings({outstr3,i2s(linkid)},";")
				end
			end
		end
	end
	
	outstr = JoinStrings({i2s(interid),i2s(numlinks),message1,message2,message3,outstr2,outstr3},",")
	
	// output for intersection check
	RunMacro("Write File", {reports.intercheck, outstr, numline, null})
	
	return(check_arr)


EndMacro

Macro "Calculate Turn Movements" (Args)
	shared thresholds, bool, ptr, labels, reports
	
	linkset = Args[1]
	headings_ib = Args[2]
	headings_ob = Args[3]
	aadt_ib = Args[4]
	aadt_ob = Args[5]
	turn_movs = Args[6]
	check_arr = Args[7]
	interid = Args[8]
	numlinks = Args[9]
	line_count = Args[10]
	
	message = "0"
	
	// total count by inbound and outbound
	aadt_ib_total = Sum(aadt_ib)
	aadt_ob_total = Sum(aadt_ob)	
	
	// 2.3: Check Summary
	if (sum(check_arr)>0 and aadt_ib_total=aadt_ob_total) then message = "1" //Flows from one or more arms are too high to calculate turning movements
	
	// check 3: turning movements
	if message="0" then do
		// find correspnding outbound direction for an inbound. for cases where an approach has a median and thus two links represent inbound and outbound directions
		for i=1 to linkset.length do
			ib_head = headings_ib[i]
			ob_head = headings_ob[i]
			if (!bool.uturn) then do // if no u-turn allowed
				if (ib_head <> null and ob_head = null) then do  // only inbound
					for j=1 to linkset.length do
						if (j<>i) then do
							ob_head = headings_ob[j]
							if (ob_head <> null and headings_ib[j]=null) then do // find the link that is only outbound
								// if difference is less than the acceptable difference in heading then consider the link as correspnding outbound link
								// and as no u-turn so set 0 movement from inbound link to the outbound link
								if abs(ib_head-ob_head)<=thresholds.heading then RunMacro("Set initial turn movements",{i,j,0})
							end
						end
					end
				end
			end
		end
		
		// iterative factoring
		iter=0
		gap=linkset.length
		while (gap > thresholds.gap and iter < 201) do
			// column factoring - outbound
			for j=1 to linkset.length do
				target = aadt_ob[j]
				trans_array = TransposeArray(turn_movs) // to get sum of columns
				total = sum(trans_array[j]) // column total
				
				// calculate ratio to factor turn movements
				if total >0 then ratio = target/total
				else ratio = 0
				
				// for each row in jth column
				for i=1 to linkset.length do 
					turn_movs[i][j] = turn_movs[i][j] * ratio 
				end
			
			end
			
			// row factoring - inbound
			for i=1 to linkset.length do
				target = aadt_ib[i]
				total = sum(turn_movs[i])  //row total
				
				// calculate ratio to factor turn movements
				if (total >0) then ratio = target/total
				else ratio = 0
				
				// for each row in jth column
				for j=1 to linkset.length do 
					turn_movs[i][j] = turn_movs[i][j] * ratio
				end
			
			end

			// calculate column gap
			gap=0
			for j=1 to linkset.length do
				target = aadt_ob[j]
				trans_array = TransposeArray(turn_movs) // to get sum of columns
				total = sum(trans_array[j]) // column total
				
				// calculate ratio to factor turn movements
				if (total >0) then ratio = target/total
				else ratio = 0
				
				gap = gap + abs(1-ratio)
			end
			
			iter=iter+1
			message = "0" //Completed
			if (iter>200) then message = "2" //Turning movements cannot be calculated, please check input flows
		end
	end
	
	// reporting
	if message = "0" then do
		outstr = JoinStrings({i2s(interid), i2s(numlinks), message, "Inbound/Outbound"},",")
		for i=1 to linkset.length do
			outstr = JoinStrings({outstr,i2s(linkset[i])},",")
		end
		outstr = JoinStrings({outstr,"TOTAL"},",")
		//write first line for an intersection
		RunMacro("Write File",{reports.interturns, outstr, line_count, numlinks})
		
		outstr_obtotal = ",,,TOTAL"
		// write turning movements
		for i=1 to linkset.length do
			outstr_obtotal = JoinStrings({outstr_obtotal, r2s(aadt_ob[i])},",")
			outstr = ",,," + i2s(linkset[i])
			for j=1 to linkset.length do
				outstr = JoinStrings({outstr,r2s(turn_movs[i][j])},",")
			end
			outstr = JoinStrings({outstr,r2s(aadt_ib[i])},",")
			line_count = line_count + 1
			RunMacro("Write File",{reports.interturns, outstr, line_count, numlinks})
		end
		line_count = line_count + 1
		RunMacro("Write File",{reports.interturns, outstr_obtotal, line_count, numlinks})
	end
	else do 
		outstr = JoinStrings({i2s(interid), i2s(numlinks), message},",")
		RunMacro("Write File",{reports.interturns, outstr, line_count, 0})
	end
	
	return(line_count)
	
EndMacro

Macro "Set initial turn movements" (Args)
	shared link_set, turn_movements
	shared bool
	
	// number of  movements (links)
	numlinks = link_set.length
	
	// passed values
	idx_ib = Args[1]
	idx_ob = Args[2]
	initial_value = Args[3]
	
	// inbound and outbound indices
	if idx_ib >0 and idx_ob>0 then do
		turn_movements[idx_ib][idx_ob] = initial_value
	end
	// inbound turn movements
	else if idx_ib >0 then do
		for j=1 to numlinks do
			turn_movements[idx_ib][j] = initial_value  // set all turn movements to initial value
			if (!bool.uturn and idx_ib=j) then turn_movements[idx_ib][j] = 0  // if no uturn, set diagonal column to 0
		end
	end
	// outbound turn movements
	else if idx_ob > 0 then do
		for i=1 to numlinks do 
			turn_movements[i][idx_ob] = initial_value  // set all turn movements to initial value
			if (!bool.uturn and i=idx_ob) then turn_movements[i][idx_ob] = 0 // if no uturn, set diagonal column to 0
		end
	
	end

EndMacro

Macro "Intersection AADT Calculations" (Args)
	shared ptr, thresholds, reports
	
	linkset = Args[1]
	miss_ib = Args[2]
	miss_ob = Args[3]
	aadt_ib = Args[4]
	aadt_ob = Args[5]
	interid = Args[6]
	numlinks = Args[7]
	numline = Args[8]
	
	if interid = 110663 then do
		test=99
	end	
	
	// total counts by inbound and outbound
	aadt_ib_total = Sum(aadt_ib)
	aadt_ob_total = Sum(aadt_ob)

	aadt_miss_ib= sum(miss_ib)
	aadt_miss_ob = sum(miss_ob)
	
	total_missing= aadt_miss_ib+aadt_miss_ob
	
	// position of the missing AADT in link set
	pos_inbound = ArrayPosition(miss_ib,{1},)
	pos_outbound = ArrayPosition(miss_ob,{1},)
	
	// linkid of the link with missing count
	if pos_inbound>0 then linkid_inbound = linkset[pos_inbound]
	if pos_outbound>0 then linkid_outbound = linkset[pos_outbound]
	
	// missing only one inbound count
	if aadt_miss_ib=1 and aadt_miss_ob=0 then do
		message = "1" // calculated a missing inbound count
		new_aadt_inbound = aadt_ob_total - aadt_ib_total
		if new_aadt_inbound < 0 then message = "2" // Cannot calculate a missing inbound count –total outbound is less than total inbound
	end
	
	// missing only one outbound count
	else if aadt_miss_ib=0 and aadt_miss_ob=1 then do
		message = "3"
		new_aadt_outbound = aadt_ib_total - aadt_ob_total
		if new_aadt_outbound < 0 then message = "4" // Cannot calculate a missing outbound count –total inbound is lower than total outbound
	end
	
	// missing one approach counts (both inbound and outbound)
	else if aadt_miss_ib=1 and aadt_miss_ob=1 then do
		message = "5" //Calculated missing approach counts (both inbound and outbound)
		// range of inbound aadt
		if aadt_ob_total > 0 then do
			new_aadt_inbound_low = thresholds.countlow * aadt_ob_total
			new_aadt_inbound_high = thresholds.counthigh * aadt_ob_total
		end
		else message = "6" // Cannot calculate - outbount flow is not available
		if aadt_ib_total > 0 then do
			// range of outbound aadt
			new_aadt_outbound_low = thresholds.countlow * aadt_ib_total
			new_aadt_outbound_high = thresholds.counthigh * aadt_ib_total
		end
		else message = "6" // Cannot calculate - inbound flow is not available
		
	end
	
	outstr1=" "
	outstr2=" "
	// for output
	if total_missing = 1 then do
		// write counts in low range column
		if aadt_miss_ib=1 then do
			// inbound count
			outstr1 = JoinStrings({"Link (", i2s(linkid_inbound), ") AADT count: ",r2s(new_aadt_inbound)},"")
			if message = "2" then outstr1 = JoinStrings({"Link (", i2s(linkid_inbound), ")"},"")

		end
		else do
			// outbound count
			outstr2 = JoinStrings({"Link (", i2s(linkid_outbound), ") AADT count: ",r2s(new_aadt_outbound)},"")
			if message = "4" then outstr2 = JoinStrings({"Link (", i2s(linkid_outbound), ")"},"")
		end
		
	end
	else do
		// missing inbound and outbound

		if aadt_ob_total>0 then outstr1 = JoinStrings({"Link (", i2s(linkid_inbound),") AADT count range: (",r2s(new_aadt_inbound_low),"-",r2s(new_aadt_inbound_high),")"},"")
		else outstr1 = JoinStrings({"Link (", i2s(linkid_inbound), ")"},"")

		if aadt_ib_total>0 then outstr2 = JoinStrings({"Link (", i2s(linkid_outbound), ") AADT count range: (",r2s(new_aadt_outbound_low),"-",r2s(new_aadt_outbound_high),")"},"")
		else outstr2 = JoinStrings({"Link (", i2s(linkid_outbound), ")"},"")

/*												
		// missing inbound and outbound
		if linkid_inbound = linkid_outbound then do
			// write in one line
			outstr = JoinStrings({"linkid: ", i2s(linkid_inbound), 
								". Inbound AADT count range: (",r2s(new_aadt_inbound_low),",",r2s(new_aadt_inbound_high), 
								") and Outbound AADT count range: (",r2s(new_aadt_outbound_low),",",r2s(new_aadt_outbound_high),")"},"")
								
		end
		else do
			outstr = JoinStrings({"Inbound link (", i2s(linkid_inbound), 
								") AADT count range: (",r2s(new_aadt_inbound_low),",",r2s(new_aadt_inbound_high), 
								"). Outbound link (", i2s(linkid_outbound), ") AADT count range: (",r2s(new_aadt_outbound_low),",",r2s(new_aadt_outbound_high),")"},"")
			//outstr = "different approaches"
		end
*/						
	end
	
	outstr = JoinStrings({i2s(interid),i2s(numlinks),message, outstr1, outstr2},",")
	RunMacro("Write File",{reports.intercalc, outstr, numline, null})

EndMacro

Macro "Intersection Missing More" (Args)
	shared ptr, reports
	
	linkset = Args[1]
	miss_ib = Args[2]
	miss_ob = Args[3]
	interid = Args[4]
	numlinks = Args[5]
	
	outstr1 = ""
	outstr2 = ""
	for i=1 to linkset.length do 
		if miss_ib[i] = 1 then do
			if outstr1="" then outstr1 = JoinStrings({outstr1, linkset[i]},"")
			else outstr1 = JoinStrings({outstr1, linkset[i]},";")
		end
		if miss_ob[i] = 1 then do 
			if outstr2="" then outstr2 = JoinStrings({outstr2, linkset[i]},"")
			else outstr2 = JoinStrings({outstr2, linkset[i]},";")
		end
	end
	
	if outstr1="" then outstr1=" "
	if outstr2="" then outstr2=" "
	
	//write output
	outstr = JoinStrings({i2s(interid),i2s(numlinks),outstr1, outstr2},",")
	RunMacro("Write File",{reports.intermissing, outstr, null, null})

EndMacro

Macro "Make Selection" (Args)
	shared inputs, outputs, fields
	
	query2 = Args[1]

	view_name = Args[2]
	
//	if inputs.query=null then query=JoinStrings({'Select * where ',fields.linkid,">0",query2},"")  // if null then select all links, ID>0
	if inputs.query=null then query=JoinStrings({'Select * where ',fields.linkid,">0"},"")  // query2 may be invalid check for centroids in other models
	else query = JoinStrings({'Select * where ',inputs.query, query2},"")
	
	// verify query
	VerifyQuery(query)
	
	count = SelectByQuery(view_name, "Several", query, )
	
	Return(count)

EndMacro

Macro "Open File and Write Header" (Args)
	shared outputs, reports, labels
	
	ftype = Args[1]
	
	// turning movements output
	if ftype = reports.interturns then do
		sfile = OpenFile(outputs.interturns,"w")
		WriteLine(sfile, JoinStrings({"interid", "linkcount" ,"msg", "turn_movements", "link1", "link2",  "link3","link4", "link5", "link6", "msg_value", "label"},","))				   
	end
	// intersection flow conservation checks
	else if ftype = reports.intercheck then do
		sfile = OpenFile(outputs.intercheck,"w")
		WriteLine(sfile, JoinStrings({"interid", "linkcount" ,"msg_check1" ,"msg_check2", "msg_check3", "linkid_check2", "linkid_check3", "msg_value", "label"},","))
	end
	// intersection with AADT calculations
	else if ftype = reports.intercalc then do
		sfile = OpenFile(outputs.intercalc,"w")
		WriteLine(sfile, JoinStrings({"interid", "linkcount", "msg", "inbound", "outbound", "msg_value", "label"},","))			
	end
	// count vs capacity
	else if ftype = reports.capcheck then do
		sfile = OpenFile(outputs.capcheck,"w")
		WriteLine(sfile, JoinStrings({"linkid", "dir", "funcclass", "countab", "countba", "capab", "capba", "ratio_ab", "ratio_ba", "msg_ab", "msg_ba", "msg_value", "label"},","))
	end
	// intersections withcount missing on more than one approach
	else if ftype = reports.intermissing then do
		sfile = OpenFile(outputs.intermissing,"w")
		WriteLine(sfile, JoinStrings({"interid","linkcount" ,"inboundlinks","outboundlinks"},","))
	end
	// count coverage summary
	else if ftype = reports.countcoverage then do
		sfile = OpenFile(outputs.countcoverage,"w")
		WriteLine(sfile, "Tennessee Network Count Coverage Report\n")
	end
	// propagated counts
	else if ftype = reports.propagatedcounts then do
		sfile = OpenFile(outputs.propagatedcounts,"w")
		WriteLine(sfile, JoinStrings({"linkid", "direction", "countAB", "countBA", "message", "message_value", "label"},","))	
	end
	// summary of all reports
	else if ftype = reports.metafile then do
		sfile = OpenFile(outputs.metafile,"w")
		WriteLine(sfile, "Summary of reports generated by the tool\n")
	end
	
	return(sfile)

EndMacro

Macro "Write File" (Args)
	shared ptr, reports, labels
	
	ftype = Args[1]
	outstr = Args[2]
	num_line = Args[3]
	num_links = Args[4]
	
	if ftype = reports.interturns then do 
		label_info = labels.interturns
		writer = ptr.interturns
	end
	else if ftype = reports.capcheck then do
		label_info = labels.capcheck
		writer = ptr.capcheck
	end
	else if ftype = reports.intercalc then do
		label_info = labels.intercalc
		writer = ptr.intercalc
	end
	else if ftype = reports.intercheck then do
		label_info = labels.intercheck
		writer = ptr.intercheck
	end
	else if ftype = reports.propagatedcounts then do
		label_info = labels.propagatedcounts
		writer = ptr.propagatedcounts
	end
	else if ftype = reports.intermissing then writer = ptr.intermissing
	else if ftype = reports.countcoverage then writer = ptr.countcoverage
	
	if num_line <> null and num_line <= label_info.len then do
		if num_links = null then outstr = JoinStrings({outstr,label_info.text[num_line]},",")
		else if num_links=0 then outstr = JoinStrings({outstr,",,,,,,",label_info.text[num_line]},",")
		else if num_links=3 then outstr = JoinStrings({outstr,",",label_info.text[num_line]},",")
		else if num_links=4 then outstr = JoinStrings({outstr,"",label_info.text[num_line]},",")
		else outstr = JoinStrings({outstr,label_info.text[num_line]},",")
	end
	
	WriteLine(writer, outstr)
	
EndMacro

/*
Count coverage statistics to report:

1. % of links with counts - network and by functional class
2. % of link lane miles with counts - network and by functional class
3. Avg. daily capacity of links with counts - network and by functional class
4. Avg. daily capacity of links without counts - network and by functional class

Intersection level:
4. No. of 3-way junctions with counts on all approaches
5. No. of 3-way junctions with counts on all but one approach
6. No. of 4-way junctions with counts on all approaches
7. No. of 4-way junctions with counts on all but one approach
8. No. of 5-way junctions with counts on all approaches
9. No. of 5-way junctions with counts on all but one approach

ODME: 
10. Avg. number of OD shortest paths on links with counts
11. Avg. number of OD shortest paths on links without counts

*/

Macro "Count Coverage"
	shared outputs, fields, query
	shared links_array, layers, db_layers, region, link_stats
	shared intertype_array, ptr
	shared withcounts, withoutcounts
	
	ptr.countcoverage = RunMacro("Open File and Write Header",{reports.countcoverage})
	ptr.countcoverage = OpenFile(outputs.countcoverage,"w")
	ok=0
	// add fields for new counts
	fields.count_AB_new = fields.count_AB+"_new"
	fields.count_BA_new = fields.count_BA+"_new"
	fields.msg = "msg"
	
	countfields_new = {{fields.count_AB_new}, {fields.count_BA_new}, {fields.msg}}

	SetLayer(layers.link)
	vw = GetView()
	on notfound goto AddNewFields
	
	fieldname = vw+"."+fields.count_AB_new
	GetField(fieldname)
	goto CalculateFields
	
AddNewFields:

	// add fields to network
	strct = GetTableStructure(vw)
	for i=1 to strct.length do
		strct[i] = strct[i] + {strct[i][1]}
	end

	strct = strct + {{fields.count_AB_new, "Real", 14, 2, "True", , , , , , , null}}
	strct = strct + {{fields.count_BA_new, "Real", 14, 2, "True", , , , , , , null}}
	strct = strct + {{fields.msg, "Int", 14, 0, "True", , ,"0-No counts; 1-Existing counts; 2-Propagated counts; 3-Counts mismatched" , , , , null}}
	
	ModifyTable(view1,strct)
	
CalculateFields:
	on error, notfound goto quit
	output_fields = {{"countAB"},{"countBA"},{"message"}}
	
	RunMacro("TCB Init")
	
	for i=1 to countfields_new.length do
		Opts = null
		Opts.Input.[Dataview Set] = {{db_layers.link, outputs.propagatedcounts,{"ID"},{"linkid"}},"JoinedView"}
		Opts.Global.Fields = countfields_new[i]
		Opts.Global.Method = "Formula"
		Opts.Global.Parameter = output_fields[i]
		ok = RunMacro("TCB Run Operation", 1, "Fill Dataview", Opts, &Ret)
		//if !ok then goto quit
	end
	
	// -------------------------------------- SUMMARY -----------------------------------
	stime = GetDateAndTime()
	WriteLine(ptr.countcoverage, "\nStarted Summary - " +SubString(stime,1,3)+","+SubString(stime,4,7)+""+SubString(stime,20,5)+" ("+SubString(stime,12,8)+") ")
	
	// Summary with existing counts
	query2 = " and (msg=1)"   // existing counts
	withcounts = RunMacro("Network Summary", {"LinksWithCounts", query2, "Existing"})	
	
	query2 = " and (msg=0 or msg>1)"  // existing no counts
	withoutcounts = RunMacro("Network Summary", {"LinksWithoutCounts", query2, "Existing"})	
	
	WriteLine(ptr.countcoverage, "\n\nNETWORK SUMMARY")
	WriteLine(ptr.countcoverage, "\nExisting Counts")
	RunMacro("Write Network Summary")
	
	// Summary after couns propagation
	query2 = " and (msg>0)" // links with counts
	withcounts = RunMacro("Network Summary", {"LinksWithCounts", query2, "Propagated"})
	
	query2 = " and (msg=0)"	// links without counts
	withoutcounts = RunMacro("Network Summary", {"LinksWithoutCounts", query2, "Propagated"})
	
	WriteLine(ptr.countcoverage, "\nPropagated Counts")
	RunMacro("Write Network Summary")
	
	Writeline(ptr.countcoverage, "\n\n\nSUMMARY OF LINKS WITH COUNTS (AFETR COUNT PROPAGATION)")
	RunMacro("Summary by Functional Class",{withcounts, "With Counts"})
	
	Writeline(ptr.countcoverage, "\n\n\nSUMMARY OF LINKS WITHOUT COUNTS (AFETR COUNT PROPAGATION)")
	RunMacro("Summary by Functional Class",{withoutcounts, "Without Counts"})

	// --------------------------- INTERCHANGE SUMMARY --------------------------------
	
	{num_inter3,inter3_allcounts,inter3_missingone,inter3_missingmore}={0,0,0,0}
	{num_inter4,inter4_allcounts,inter4_missingone,inter4_missingmore}={0,0,0,0}
	{num_inter5,inter5_allcounts,inter5_missingone,inter5_missingmore}={0,0,0,0}
	{num_inter2,num_inter6}={0,0}
	
	// interchange type summary
	for inter=1 to intertype_array.length do
		nlinks = intertype_array[inter][1]
		type = intertype_array[inter][2]
		
		if nlinks=2 then num_inter2 = num_inter2+1
		// 3-way junction
		else if nlinks=3 then do
			num_inter3 = num_inter3+1
			if type = 0 then inter3_allcounts=inter3_allcounts+1          // counts on all approaches
			else if type=1 then inter3_missingone=inter3_missingone+1     // counts on all but one approach
			else inter3_missingmore=inter3_missingmore+1                  // more than one approach counts are missing
		end
		// 4-way junction
		else if nlinks=4 then do
			num_inter4 = num_inter4+1
			if type=0 then inter4_allcounts=inter4_allcounts+1
			else if type=1 then inter4_missingone=inter4_missingone+1
			else inter4_missingmore=inter4_missingmore+1
		end
		// 5-way junction
		else if nlinks=5 then do
			num_inter5 = num_inter5+1
			if type=0 then inter5_allcounts=inter5_allcounts+1
			else if type=1 then inter5_missingone=inter5_missingone+1
			else inter5_missingmore=inter5_missingmore+1
		end
		else num_inter6 = num_inter6+1		
	
	end
	
	WriteLine(ptr.countcoverage, "\n\n\nINTERCHANGE SUMMARY (AFETR COUNT PROPAGATION)")
	WriteLine(ptr.countcoverage, "=====================================================================================================================")
	WriteLine(ptr.countcoverage, "Type	 		CountsOnAll		(%)	   	MissingOneApproach			(%)			MissingMore			(%)	 		TOTAL")
	WriteLine(ptr.countcoverage, "=====================================================================================================================")
	//WriteLine(ptr.countcoverage, "2-way			  																   " + Format(num_inter2,",00000"))
	WriteLine(ptr.countcoverage, "3-way			  " + Format(inter3_allcounts, ",00000") + "	" + Format(100*inter3_allcounts/num_inter3, "00.00") 
													+ "%					" + Format(inter3_missingone,",,00000") + "		" + Format(100*inter3_missingone/num_inter3, "00.00") 
													+ "%				 " + Format(inter3_missingmore,",00000") + "		" + Format(100*inter3_missingmore/num_inter3, "00.00") 
													+ "%	 	   " + Format(num_inter3,",00000"))
													
	WriteLine(ptr.countcoverage, "4-way			  " + Format(inter4_allcounts, ",00000") + "	" + Format(100*inter4_allcounts/num_inter4, "00.00") 
													+ "%					" + Format(inter4_missingone,",,00000") + "		" + Format(100*inter4_missingone/num_inter4, "00.00") 
													+ "%				 " + Format(inter4_missingmore,",00000") + "		" + Format(100*inter4_missingmore/num_inter4, "00.00") 
													+ "%	 	   " + Format(num_inter4,",00000"))
	WriteLine(ptr.countcoverage, "5-way			  " + Format(inter5_allcounts, ",00000") + "	" + Format(100*inter5_allcounts/num_inter5, "00.00") 
													+ "%					" + Format(inter5_missingone,",,00000") + "		" + Format(100*inter5_missingone/num_inter5, "00.00") 
													+ "%				 " + Format(inter5_missingmore,",00000") + "		" + Format(100*inter5_missingmore/num_inter5, "00.00") 
													+ "%	 	   " + Format(num_inter5,",00000"))
	//WriteLine(ptr.countcoverage, "6-way			  																   " + Format(num_inter6,",00000"))
	
	total_num = num_inter2+num_inter3+num_inter4+num_inter5+num_inter6
	total_allcounts = inter3_allcounts+inter4_allcounts+inter5_allcounts
	total_missingone = inter3_missingone+inter4_missingone+inter5_missingone
	total_missingmore = inter3_missingmore+inter4_missingmore+inter5_missingmore
/*	
	WriteLine(ptr.countcoverage, "=========================================================================================")
	WriteLine(ptr.countcoverage, "TOTAL			  " + Format(total_allcounts, ",00000") + "						" + Format(total_missingone,",,00000") + "				 "
											+ Format(total_missingmore,",00000") + "	 	   " + Format(total_num,",000000"))
*/	
	// --------------------------------------- ODME SUMMARY ------------------------------------------
	
	RunMacro("Build Network") 			// build a network
	RunMacro("Create Demand Matrix") 	// create a demand matrix containing 1's
	RunMacro("Run AON Assignment") 		// perform all-or-nothing assginment in free-flow conditions
	
	// join the output flows to network and summarize for links with counts and without counts 
	view_flow = OpenTable("Flow Table", "FFB", {outputs.flowtable},{{"Shared", "True"}})
	SetLayer(layers.link)
	vw = GetView()
	joined_view = JoinViews("links+flows",vw + ".ID",view_flow + ".ID1", )
	
	SetView(joined_view)
	
	// select links with counts - msgcode =1 or 2 or 3 (mismatched counts)
	query2 = " and (msg>0)"
	view_name = "LinksWithCounts"
	n=RunMacro("Make Selection", {query2, view_name})
	
	// make view set
	view_set = joined_view + "|" + view_name
	
	flow_array = GetDataVector(view_set, "Tot_Flow",)
	withcount.avgflow = VectorStatistic(flow_array,"mean",)
	
	// select links without counts - msgcode =1 or 2 or 3 (mismatched counts)
	query2 = " and (msg=0)"
	view_name = "LinksWithoutCounts"
	
	n=RunMacro("Make Selection", {query2, view_name})
	
	view_set = joined_view + "|" + view_name 
	
	flow_array = GetDataVector(view_set, "Tot_Flow",)
	withoutcount.avgflow = VectorStatistic(flow_array,"mean",)	
	
	WriteLine(ptr.countcoverage, "\n\n\nODME SUMMARY (AFETR COUNT PROPAGATION)")
	
	WriteLine(ptr.countcoverage, "==========================================================")
	WriteLine(ptr.countcoverage, "	 				Avg. Number of Shortest Paths")
	WriteLine(ptr.countcoverage, "==========================================================")	
	WriteLine(ptr.countcoverage, "With Counts			  		" + Format(withcount.avgflow,",000000"))
	WriteLine(ptr.countcoverage, "Without Counts		  		" + Format(withoutcount.avgflow,",000000"))
	
	stime = GetDateAndTime()
	WriteLine(ptr.countcoverage, "\n\n\nFinished Summary - " +SubString(stime,1,3)+","+SubString(stime,4,7)+""+SubString(stime,20,5)+" ("+SubString(stime,12,8)+") ")
	
	CloseFile(ptr.countcoverage)
	CloseView(joined_view)
	
	ok=1
	
	quit:
	Return(RunMacro("TCB Closing", ok, True))

EndMacro

Macro "Summary by Functional Class" (Args)
	shared ptr, fields, outputs, link_stats
	
	sel_arr = Args[1]
	type = Args[2]
	
	field_names = {fields.funcclass,"N "+fields.linkid,fields.len,"Avg "+fields.count_AB_new,"Avg "+fields.count_BA_new, "Avg "+fields.capacity_AB,"Avg "+fields.capacity_BA}
	
	Writeline(ptr.countcoverage, "\nBy Functional Class")
	WriteLine(ptr.countcoverage, "======================================================================================================================================")
	WriteLine(ptr.countcoverage, "FUNCCLASS	    	 NumLinks	   		(%)		LaneMiles	 		(%)	  Avg_AADT_AB	  Avg_AADT_BA	    Avg_Cap_AB	    Avg_Cap_BA")
	WriteLine(ptr.countcoverage, "======================================================================================================================================")
	
	curr_view = GetView()
	// Compute Statistics by facility type
	temp_output = outputs.dir + "\\temp.csv"
	sel_arr.summary = AggregateTable("NoCount Summary", sel_arr.viewset, "CSV", temp_output, fields.funcclass,
							{{fields.linkid,"count"},
							 {fields.len,"sum"},
							 {fields.count_AB_new,"avg"},
							 {fields.count_BA_new,"avg"},
							 {fields.capacity_AB,"avg"},
							 {fields.capacity_BA,"avg"}},null)
	
	rec = GetFirstRecord(sel_arr.summary+"|", null)
	
	nocount_stats = {}
	
	while rec <> null do
		
		func_class = sel_arr.summary.(field_names[1])
		linkcount = sel_arr.summary.(field_names[2])
		lanemiles = sel_arr.summary.(field_names[3])
		avg_count_ab = sel_arr.summary.(field_names[4])
		avg_count_ba = sel_arr.summary.(field_names[5])
		avg_cap_ab = sel_arr.summary.(field_names[6])
		avg_cap_ba = sel_arr.summary.(field_names[7])

		stats = link_stats.(i2s(func_class))
		
		perclinkcount = 100*linkcount/stats[1]
		perclanemiles = 100*lanemiles/stats[2]
		if type = "With Counts" then do
			if (perclinkcount>99.99) then do
				WriteLine(ptr.countcoverage, "   " + Format(func_class,"00") +"				  " + Format(linkcount,",000000") +"	   " + Format(perclinkcount,"000.00") +"%		" + Format(lanemiles,",00000.00")
										+ "	   " + Format(perclanemiles,"000.00") + "%		" + Format(avg_count_ab,",00000.00") + "		" + Format(avg_count_ba,",000000.00")
										+ "	    " + Format(avg_cap_ab,",000000.00") + "		" + Format(avg_cap_ba,",000000.00"))
			end
			else do
				WriteLine(ptr.countcoverage, "   " + Format(func_class,"00") +"				  " + Format(linkcount,",000000") +"	    " + Format(perclinkcount,"00.00") +"%		" + Format(lanemiles,",00000.00")
										+ "	    " + Format(perclanemiles,"00.00") + "%		" + Format(avg_count_ab,",00000.00") + "		" + Format(avg_count_ba,",00000.00")
										+ "		" + Format(avg_cap_ab,",000000.00") + "		" + Format(avg_cap_ba,",000000.00"))
			end
		end
		else do
			if (perclinkcount>99.99) then do
				WriteLine(ptr.countcoverage, "   " + Format(func_class,"00") +"				  " + Format(linkcount,",000000") +"	   " + Format(perclinkcount,"000.00") +"%		" + Format(lanemiles,",00000.00")
										+ "	   " + Format(perclanemiles,"000.00") + "%			   " + Format(avg_count_ab,",00000.00") + "			   " + Format(avg_count_ba,",000000.00")
										+ "	    " + Format(avg_cap_ab,",000000.00") + "		" + Format(avg_cap_ba,",000000.00"))
			end
			else do
				WriteLine(ptr.countcoverage, "   " + Format(func_class,"00") +"				  " + Format(linkcount,",000000") +"	    " + Format(perclinkcount,"00.00") +"%		" + Format(lanemiles,",00000.00")
										+ "	    " + Format(perclanemiles,"00.00") + "%			   " + Format(avg_count_ab,",00000.00") + "			   " + Format(avg_count_ba,",00000.00")
										+ "		" + Format(avg_cap_ab,",000000.00") + "		" + Format(avg_cap_ba,",000000.00"))
			end		
		end
		
		rec = GetNextRecord(sel_arr.summary+"|", null, null)
		
	end

	WriteLine(ptr.countcoverage, "======================================================================================================================================")
	WriteLine(ptr.countcoverage, "TOTAL			 	  " + Format(sel_arr.numlinks, ",000000") + "		" + Format(sel_arr.perclinks,"00.00")+"%		" + Format(sel_arr.lanemiles,",00000.00")
						 + "		" + Format(sel_arr.perclanemiles,"00.00") + "%")

	CloseView(sel_arr.summary)
	DeleteFile(temp_output)
	DeleteFile(outputs.dir + "\\temp.CSX")
	DeleteFile(outputs.dir + "\\temp.DCC")

EndMacro

Macro "Network Summary" (Args)
	Shared fields, region
	
	view_name = Args[1]
	qry = Args[2]
	type = Args[3]
	
	sel_arr = {}
	
	sel_arr.view = view_name
	
	// Select links with counts
	sel_arr.numlinks = RunMacro("Make Selection", {qry, sel_arr.view})

	// make view set
	vw = GetView()
	sel_arr.viewset = vw + "|" + sel_arr.view
	
	if type = "Propagated" then do
		sel_arr.array = GetDataVectors(sel_arr.viewset, {fields.len, fields.count_AB_new, fields.count_BA_new, fields.capacity_AB, fields.capacity_BA},)
	end
	else if type = "Existing" then do
		sel_arr.array = GetDataVectors(sel_arr.viewset, {fields.len, fields.count_AB, fields.count_BA, fields.capacity_AB, fields.capacity_BA},)
	end
	
	sel_arr.lanemiles = VectorStatistic(sel_arr.array[1],"sum",)
	sel_arr.avgcountab = VectorStatistic(sel_arr.array[2],"mean",)
	sel_arr.avgcountba = VectorStatistic(sel_arr.array[3],"mean",)
	sel_arr.avgcapab = VectorStatistic(sel_arr.array[4],"mean",)
	sel_arr.avgcapba = VectorStatistic(sel_arr.array[5],"mean",)
	sel_arr.perclinks = 100*(sel_arr.numlinks/region.numlinks)
	sel_arr.perclanemiles = 100*sel_arr.lanemiles/region.lanemiles	
	
	return(sel_arr)
	
EndMacro

Macro "Write Network Summary"
	shared withcounts, withoutcounts, region, ptr
	
	WriteLine(ptr.countcoverage, "=====================================================================================================================================")
	WriteLine(ptr.countcoverage, "LinkType	   		 NumLinks	   		(%)		LaneMiles	 		(%)	  Avg_AADT_AB	  Avg_AADT_BA	   Avg_Cap_AB	   Avg_Cap_BA")
	WriteLine(ptr.countcoverage, "=====================================================================================================================================")	
	WriteLine(ptr.countcoverage, "With Counts	 		  " + Format(withcounts.numlinks,",000000") + "		" + Format(withcounts.perclinks,"00.00")+"%		" + Format(withcounts.lanemiles,",00000.00")
						 + "		" + Format(withcounts.perclanemiles,"00.00") + "%		" + Format(withcounts.avgcountab,",00000.00") + "		" + Format(withcounts.avgcountba,",00000.00")
							  + "		" + Format(withcounts.avgcapab,",00000.00") + "		" + Format(withcounts.avgcapba,",00000.00"))
	
	WriteLine(ptr.countcoverage, "Without Counts 		  " + Format(withoutcounts.numlinks,",000000") + "		" + Format(withoutcounts.perclinks,"00.00")+"%		" + Format(withoutcounts.lanemiles,",00000.00")
						 + "		" + Format(withoutcounts.perclanemiles,"00.00") + "%		" + Format(withoutcounts.avgcountab,",00000.00") + "		" + Format(withoutcounts.avgcountba,",00000.00")
							  + "		" + Format(withoutcounts.avgcapab,",00000.00") + "		" + Format(withoutcounts.avgcapba,",00000.00"))
	
	WriteLine(ptr.countcoverage, "=====================================================================================================================================")
	WriteLine(ptr.countcoverage, "TOTAL				  " + Format(region.numlinks, ",000000") + "	   100.00%		" + Format(region.lanemiles,",00000.00")
						 + "	   100.00%")

EndMacro

Macro "Build Network"
	shared inputs, outputs, fields
	shared db_layers, layers, query
	
	// for debug
	//query = 'Select * where STATE = "TN"'
	query=JoinStrings({'Select * where ',inputs.query},"")
	
	// verify query
	VerifyQuery(query)
	
	RunMacro("TCB Init")
	
    Opts = null
    Opts.Input.[Link Set] = {db_layers.link , layers.link, "Selection", query}
    Opts.Global.[Network Label] = "Based on "+db_layers.link
	Opts.Global.[Network Options].[Turn Penalties] = "Yes"
	Opts.Global.[Network Options].[Keep Duplicate Links] = "FALSE"
	Opts.Global.[Network Options].[Ignore Link Direction] = "FALSE"
	Opts.Global.[Network Options].[Time Units] = "Minutes"
	Opts.Global.[Link Options] = {{"Length", {layers.link+".Length", layers.link+".Length", , , "False"}}, 
	{"[AB_DlyCap / BA_DlyCap]", {layers.link+".AB_DLYCAP", layers.link+".BA_DLYCAP", , , "False"}}, 
	{"[AB_AFFSpd / BA_AFFSpd]", {layers.link+".AB_AFFSpd", layers.link+".BA_AFFSpd", , , "False"}}, 
	{"[AB_AFFTime / BA_AFFTime]", {layers.link+".AB_AFFTime", layers.link+".BA_AFFTime", , , "False"}}}
	Opts.Global.[Length Units] = "Miles"
	Opts.Global.[Time Units] = "Minutes"
	Opts.Output.[Network File] = outputs.hwynet
	ok = RunMacro("TCB Run Operation", "Build Highway Network", Opts, &Ret)
	if !ok then goto quit 

	// centroids
	Opts = null
	Opts.Input.Database = inputs.dbfile
	Opts.Input.Network = outputs.hwynet
	Opts.Input.[Centroids Set] = {inputs.dbfile+"|"+layers.node, layers.node, "Selection", "Select * where Centroid=1 and TAZID<4000"}
	ok = RunMacro("TCB Run Operation", "Highway Network Setting", Opts, &Ret)
	if !ok then goto quit 

	quit:
	Return(ok)
	//Return(RunMacro("TCB Closing", ok, True))	
	
EndMacro

Macro "Create Demand Matrix"
    shared inputs, outputs, layers
	
	query = "Select * where Centroid=1 and TAZID<4000"
	
	SetLayer(layers.node)
	n = SelectByQuery("Centroids","Several",query,)
	
	view = GetView()
	view_set = view + "|Centroids"
    
    // create a new matrix
    mat =CreateMatrix({view_set, "ID", "Rows"},
     {view_set, "ID", "Cols"},
     {{"File Name", outputs.demandmat}, {"Type", "Float"},
     {"Tables",{"Demand"}},{"Compression",1}, {"Do Not Initialize", "True"}})
    
	mc = CreateMatrixCurrency(mat, "Demand","Rows","Cols",)
	mc:=nz(mc)
	
	// set demand to 1
	FillMatrix(mc,,,{"Add",1},)
    
    // null out the matrix and currency handles
    mc_rev = null
    mat  = null
    
    // Close view
    CloseView(view)

EndMacro

Macro "Run AON Assignment"
	shared inputs, outputs
	shared db_layers, layers
	
	RunMacro("TCB Init")
	
	// perform all or nothing (AON) assignment
	Opts = null
	Opts.Input.Database = inputs.dbfile
	Opts.Input.Network = outputs.hwynet
	Opts.Input.[OD Matrix Currency] = {outputs.demandmat, "Demand", "Rows", "Cols"}
	Opts.Field.[VDF Fld Names] = {"[AB_AFFTime / BA_AFFTime]", "[AB_DlyCap / BA_DlyCap]", "None", "None", "None"}
	Opts.Global.[Load Method] = "AON"
	Opts.Field.[FF Time] = "[AB_AFFTime / BA_AFFTime]"
	Opts.Global.[Loading Multiplier] = 1
	Opts.Global.[VDF DLL] = "bpr.vdf"
	Opts.Global.[VDF Defaults] = {, , 0.15, 4, 0}
	Opts.Flag.[Do Theme] = 1
	Opts.Output.[Flow Table] = outputs.flowtable
	
	ok = RunMacro("TCB Run Procedure", "Assignment", Opts, &Ret)
	if !ok then goto quit
	
	quit:
	Return(ok)
	//Return(RunMacro("TCB Closing", ok, True))

EndMacro

Macro "Write Summary of Outputs"
	shared inputs, outputs, reports, lst_reports, ptr
	
	ptr.metafile = RunMacro("Open File and Write Header",{reports.metafile})
	
	//WriteLine(ptr.metafile, "Output Folder: " + outputs.dir + "\n")
	WriteLine(ptr.metafile, "-----------------------------------------------------------------------------------------------------------------------")
	WriteLine(ptr.metafile, "Report								Output")
	WriteLine(ptr.metafile, "-----------------------------------------------------------------------------------------------------------------------")
	WriteLine(ptr.metafile, reports.countcoverage + "						" + outputs.countcoverage)
	WriteLine(ptr.metafile, reports.propagatedcounts + "					" + outputs.propagatedcounts)
	WriteLine(ptr.metafile, reports.capcheck + "				" + outputs.capcheck)
	WriteLine(ptr.metafile, reports.intercheck + "		" + outputs.intercheck)
	WriteLine(ptr.metafile, reports.interturns + "		" + outputs.interturns)
	WriteLine(ptr.metafile, reports.intercalc + "		" + outputs.intercalc)
	WriteLine(ptr.metafile, reports.intermissing + "			" + outputs.intermissing)
	WriteLine(ptr.metafile, "-----------------------------------------------------------------------------------------------------------------------")
	
	CloseFile(ptr.metafile)
	
EndMacro

Macro "Open Report" (Args)
	shared outputs, reports
			
	rprt = Args[1]
	
	RunMacro("Set output settings")
	
	if rprt = reports.metafile then LaunchDocument(outputs.metafile,)
	else if rprt = reports.countcoverage then LaunchDocument(outputs.countcoverage,)
	else if rprt = reports.capcheck then LaunchDocument(outputs.capcheck,)
	else if rprt = reports.propagatedcounts then LaunchDocument(outputs.propagatedcounts,)
	else if rprt = reports.intercheck then LaunchDocument(outputs.intercheck,)
	else if rprt = reports.interturns then LaunchDocument(outputs.interturns,)
	else if rprt = reports.intercalc then LaunchDocument(outputs.intercalc,)
	else if rprt = reports.intermissing then LaunchDocument(outputs.intermissing,)

EndMacro

Macro "Display Result" (result, result_field)
	shared inputs, outputs, fields, layers, region, reports
	shared cc_Colors

/*	
	view_names = GetViewNames()
	
	for vw=1 to view_names.length do
		CloseView(view_names[vw])
	end
*/	
	// close an existing map
	maps = GetMapNames()
	for i=1 to maps.length do
		CloseMap(maps[i])
	end
	
	// Get scope
	info = GetDBInfo(inputs.dbfile)
	network_scope = info[1]
	
	map = CreateMap("Display Result: " + result,{{"Scope", network_scope},{"Auto Project", "True"},{"Position", 0,0}})
	
	// add link and node layers
	node_layer = AddLayer(map,"nodes",inputs.dbfile, layers.node)
	link_layer = AddLayer(map,"links",inputs.dbfile, layers.link)
	
	RunMacro("Set output settings")
	RunMacro("Select Region", link_layer)
	RunMacro("Select Interchanges", node_layer)
	
	if (result = reports.propagatedcounts) or (result = reports.capcheck) then do
		SetLayer(link_layer)
		
		SetLineColor(link_layer+"|", cc_Colors.Black)
		SetLineWidth(link_layer+"|", 1.0)				
		
		RunMacro("Create Color Theme",link_layer, result, "linkid" ,result_field)
	end
	else if result = "ODME Shortest Paths" then do
		SetLayer(link_layer)
		
		RunMacro("Create Scaled Theme", link_layer, result, "ID1" , result_field)
	end
	else do
		SetLayer(node_layer)
		SetIcon(node_layer+"|","Font Character","Caliper Cartographic|4",1)
		SetIconColor(node_layer+"|",cc_Colors.Black)		
		
		RunMacro("Create Color Theme",node_layer, result, "interid" , result_field)
	end

EndMacro

Macro "Create Color Theme" (lyr, rslt,rslt_id, fld)
	shared inputs, outputs, fields, layers, region, reports
	shared cc_Colors
	
	
	// delete header file - header file sometimes causes issues if previous outputs are different
	RunMacro("Delete Header File", rslt)	
	
	// theme settings
	theme_colors_array = {cc_Colors.White, cc_Colors.Black, cc_Colors.Blue, cc_Colors.Green, cc_Colors.Red, cc_Colors.Purple, cc_Colors.Orange}
	theme_widths_array = {0.25, 1.0, 1.5, 1.5, 1.5, 1.5, 1.5}	
	theme_icons_array = {{"Font Character","Caliper Cartographic|10",0},
						 {"Font Character","Caliper Cartographic|10",5},
						 {"Font Character","Caliper Cartographic|10",5},
						 {"Font Character","Caliper Cartographic|10",5},
						 {"Font Character","Caliper Cartographic|10",5},
						 {"Font Character","Caliper Cartographic|10",5},
						 {"Font Character","Caliper Cartographic|10",5}}
	map_scope = region.nodeview
	
	// create a joined view
	if rslt = reports.propagatedcounts then do
		joined_view = lyr
		map_scope = region.view
	end
	else do
		if rslt = reports.capcheck then do
			rslt_file = outputs.capcheck
			map_scope = region.view
			RunMacro("TCB Init")
		end
		else if rslt = reports.interturns then rslt_file = outputs.interturns
		else if rslt = reports.intermissing then rslt_file = outputs.intermissing
		else if rslt = reports.intercalc then rslt_file = outputs.intercalc
		else if rslt = reports.intercheck then rslt_file = outputs.intercheck
		
		// open result table and join to layer
		result_view = OpenTable(rslt, "CSV",{rslt_file, null},{{"Shared","True"}})	
		
		//debug
		//save as binary file
		//rslt_file=outputs.dir+"\\capchecks.bin"
		//ExportView(result_view+"|","FFB",rslt_file,,)
		//CloseView(result_view)
		
		//result_view = OpenTable(rslt, "FFB",{rslt_file, null},{{"Shared","True"}})
		
		joined_view = JoinViews("links+result", lyr + ".ID", result_view + "." + rslt_id, )

	end
	
	if rslt = reports.intermissing then do

		query = 'Select *where ' + fld + '> 0'
		n=SelectByQuery(rslt, "Several", query)
		if (n>0) then do 
			SetmapScope(,GetSetScope(rslt))
			SetDisplayStatus(rslt,"Active")
			SetIcon(rslt,"Font Character","Caliper Cartographic|8",1)
			SetIconColor(rslt,cc_Colors.Red)
		end	
	
	end
	else do
		// create theme
		theme = CreateTheme(rslt,joined_view+"."+fld,"Categories",10,)
		ShowTheme(,theme)
		
		// set theme settings
		if map_scope = region.view then do
			// line layer
			SetThemeLineColors(theme, theme_colors_array)
			SetThemeLineWidths(theme, theme_widths_array)
		end
		else do
			// node layer
			SetThemeIconColors(theme, theme_colors_array)
			SetThemeIcons(theme, theme_icons_array)
		end
		
		//theme labels
		theme_values = GetThemeClassLabels(theme)
		theme_labels = RunMacro("Get Theme Labels", rslt, theme_values)
		
		SetThemeClasslabels(theme,theme_labels)
		
		SetmapScope(,GetSetScope(map_scope))
		RunMacro("G30 create legend", "Theme")
	end	
	
EndMacro

Macro "Delete Header File" (ftype)
	shared outputs
	shared reports
	
	on Error, NotFound goto quit 
	
	if ftype = reports.propagatedcounts then filepath = outputs.propagatedcounts
	else if ftype = reports.capcheck then filepath = outputs.capcheck
	else if ftype = reports.intercheck then filepath = outputs.intercheck
	else if ftype = reports.interturns then filepath = outputs.interturns
	else if ftype = reports.intercalc then filepath = outputs.intercalc
	
	file_info = SplitPath(filepath)
	
	filename_noext = file_info[3]
	header_filename = filename_noext + ".DCC"
	header_filepath = outputs.dir + "\\" + header_filename
	
	DeleteFile(header_filepath)
	
	quit:
	// file doesn't exit or being used by another program
	return(0)
	
EndMacro

Macro "Get Theme Labels" (ftype, theme_vals)
	shared reports, labels
	
	// initializae 
	theme_lbls = theme_vals
	
	if ftype = reports.intercalc then do
		// intesection AADT calculation labels start from 1 to 6
		labels_text = labels.intercalc.text
		for i = 1 to theme_vals.length do
			if theme_vals[i] = "Other" then do
				theme_lbls[i] = " "
			end
			else do
				index = StringToInt(theme_vals[i])
				theme_lbls[i] = labels_text[index]
			end
		end		
	end
	else do
		// these label values start at 0, so add 1 to index
		if ftype = reports.propagatedcounts then labels_text = labels.propagatedcounts.text
		else if ftype = reports.capcheck then labels_text = labels.capcheck.text
		else if ftype = reports.intercheck then labels_text = labels.intercheck.text
		else if ftype = reports.interturns then labels_text = labels.interturns.text	

		for i = 1 to theme_vals.length do
			if theme_vals[i] = "Other" then do
				theme_lbls[i] = " "
			end
			else do
				index = StringToInt(theme_vals[i])
				theme_lbls[i] = labels_text[index+1]
			end
		end
	end
	
	return(theme_lbls)
EndMacro

Macro "Create Scaled Theme" (lyr, rslt, rslt_id, fld)
	shared inputs, outputs, fields, layers, region, reports
	shared cc_Colors
	
	rslt_file = outputs.flowtable
	
	//open file
	result_view = OpenTable("result", "FFB", {rslt_file, null}, {{"Shared", "True"}})
	
	joined_view = JoinViews("links+result", lyr + ".ID", result_view + "." + rslt_id,)
	theme = CreateContinuousTheme(rslt, {joined_view+"." + fld},{{"Title",rslt},{"Minimum value", 1}, {"Minimum size", 1}})
	ShowTheme(,theme)
	
	SetmapScope(,GetSetScope(region.view))
	RunMacro("G30 create legend", "Theme")		
	
EndMacro
