/*------------------------------------------------------------------------------*
 * File Name:				 													*
 * Creation: 																	*
 * Purpose: OriginC Source C file												*
 * Copyright (c) ABCD Corp.	2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010		*
 * All Rights Reserved															*
 * 																				*
 * Modification Log:															*
 *------------------------------------------------------------------------------*/
 
////////////////////////////////////////////////////////////////////////////////////
// Including the system header file Origin.h should be sufficient for most Origin
// applications and is recommended. Origin.h includes many of the most common system
// header files and is automatically pre-compiled when Origin runs the first time.
// Programs including Origin.h subsequently compile much more quickly as long as
// the size and number of other included header files is minimized. All NAG header
// files are now included in Origin.h and no longer need be separately included.
//
// Right-click on the line below and select 'Open "Origin.h"' to open the Origin.h
// system header file.
#include <Origin.h>
////////////////////////////////////////////////////////////////////////////////////

//#pragma labtalk(0) // to disable OC functions for LT calling.

////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.
#include <..\Originlab\graph_utils.h>
#include <..\originlab\fu_utils.h>
#include "DataPlotter.h"
#include <math.h>

////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.
Axis DataPlotter::get_axis(int nAxisType)
{
	Axis axis;
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		axis = gl.XAxis;
		break;
	case AXIS_LEFT:
		axis = gl.YAxis;
		break;
	case AXIS_TOP:
		axis = gl.XAxis;
		break;
	case AXIS_RIGHT:
		axis = gl.YAxis;
		break;
	}
	return axis;
}
void DataPlotter::upate_axis(Tree tr, Axis axis)
{
	int nErr = axis.UpdateThemeIDs(tr.Root);// allows a tree with no theme ID to be used in ApplyFormat
	if(nErr ==0)
		axis.ApplyFormat(tr,true,true);
}
DataPlotter::DataPlotter(string sGraphPageName)
{
	GraphPage temp;
	gp = temp;
	gp.Create();
	gp.SetName(sGraphPageName);
	gl = gp.Layers(0);	
}
void DataPlotter::select_layer(int nLayerInd)
{
	if(nLayerInd < gp.Layers.Count())
		gl = gp.Layers(nLayerInd);
	else
		printf("Layer index out of range");
}
void DataPlotter::select_plot(int nPlotInd)
{
	if(nPlotInd < gl.DataPlots.Count())
		dp = gl.DataPlots(nPlotInd);
	else
		printf("Data Plot index out of range");
}

void DataPlotter::add_layer(int nAxisType, int nLinkTo = 0, int nXAxisLink = LINK_STRAIGHT, int nYAxisLink = 0)
{
	// nLinkTo will determine which axis will be visible. 
	// int nLinkTo is the index of the layer to link to
	// int nXAxisLink is how the x axis should be linked. Options are: 0, LINK_STRAIGHT, LINKED_AXIS_CUSTOM, LINKED_AXIS_ALIGN
	// int nYAxisLink is how the y axis should be linked. Options are: 0, LINK_STRAIGHT, LINKED_AXIS_CUSTOM, LINKED_AXIS_ALIGN
	bool bBottom = false, bLeft = false, bTop = false, bRight = false;
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		bBottom = true;
		break;
	case AXIS_LEFT:
		bLeft = true;
		break;
	case AXIS_TOP:
		bTop = true;
		break;
	case AXIS_RIGHT:
		bRight = true;
		break;
	}
	bool bActivateNewLayer = false;
	//add new layer
	page_add_layer(gp, bBottom, bLeft, bTop, bRight,ADD_LAYER_INIT_SIZE_POS_MOVE_OFFSET, bActivateNewLayer, nLinkTo, nXAxisLink, nYAxisLink);
	//get the newest layer index in the graph page
	int nLayerInt = gp.Layers.Count()-1;
	//store the newest layer
	gl = gp.Layers(nLayerInt);
	
}
void DataPlotter::graphpage_resize(double dWidth, double dHeight)
{
	/*
	tr.Root.Dimension.Width.dVal = dWidth;
	tr.Root.Dimension.Height.dVal = dHeight;
	
	int nErr = gp.UpdateThemeIDs(tr.Root);// allows a tree with no theme ID to be used in ApplyFormat
	if(nErr ==0)
		gp.ApplyFormat(tr,true,true);
	*/
}
void DataPlotter::show_axis(int nAxisType = AXIS_LEFT, bool bAxisOn = true, bool bLabels = true, bool bTitleOn = true, int nMajorTicks = TICK_OUT, int nMinorTicks = TICK_OUT)
{
	Tree tr;
	Axis axis = get_axis(nAxisType);
	//tr = gl.XAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		tr.Root.Ticks.BottomTicks.AddNode("Show").nVal = bAxisOn ? 1:0;
		tr.Root.Ticks.BottomTicks.LineShow.nVal = bAxisOn ? 1:0;
		tr.Root.Labels.BottomLabels.AddNode("Show").nVal = bLabels ? 1:0;
		tr.Root.Titles.BottomTitle.AddNode("Show").nVal = bTitleOn ? 1:0;
		break;
	case AXIS_LEFT:
		tr.Root.Ticks.LeftTicks.AddNode("Show").nVal = bAxisOn ? 1:0;
		tr.Root.Ticks.LeftTicks.LineShow.nVal = bAxisOn ? 1:0;
		tr.Root.Labels.LeftLabels.AddNode("Show").nVal = bLabels ? 1:0;
		tr.Root.Titles.LeftTitle.AddNode("Show").nVal = bTitleOn ? 1:0;
		break;
	case AXIS_TOP:
		tr.Root.Ticks.TopTicks.AddNode("Show").nVal = bAxisOn ? 1:0;
		tr.Root.Ticks.TopTicks.LineShow.nVal = bAxisOn ? 1:0;
		tr.Root.Labels.TopLabels.AddNode("Show").nVal = bLabels ? 1:0;
		tr.Root.Titles.TopTitle.AddNode("Show").nVal = bTitleOn ? 1:0;
		break;
	case AXIS_RIGHT:
		tr.Root.Ticks.RightTicks.AddNode("Show").nVal = bAxisOn ? 1:0;
		tr.Root.Ticks.RightTicks.LineShow.nVal = bAxisOn ? 1:0;
		tr.Root.Labels.RightLabels.AddNode("Show").nVal = bLabels ? 1:0;
		tr.Root.Titles.RightTitle.AddNode("Show").nVal = bTitleOn ? 1:0;
		break;
	}
	upate_axis(tr, axis);
	//tr = gl.YAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
}
/*
void DataPlotter::show_axis(int nAxisType = AXIS_LEFT, bool bAxisOn = true, bool bLabels = true, bool bTitleOn = true, int nMajorTicks = TICK_OUT, int nMinorTicks = TICK_OUT)
{
	//nAxisType: AXIS_LEFT, AXIS_RIGHT, AXIS_TOP, AXIS_BOTTOM
	vector<int> temp(4);
	
	temp[0] = 0;
	temp[1] = 0;
	temp[2] = 0;
	temp[3] = 0;
	
	// Show all axes, labels, and titles. 0 or 1, 1 for show.
	vector<int> vnAxes;
	vector<int> vnLabels;
	vector<int> vnTitles;
	vnAxes = temp;
	vnLabels = temp;
	vnTitles = temp;
	// Set the major tick and minor tick of all axes as IN format
	temp[0] = TICK_IN;
	temp[1] = TICK_IN;
	temp[2] = TICK_IN;
	temp[3] = TICK_IN;
	vector<int> vnMajorTicks;
	vector<int> vnMinorTicks;
	vnMajorTicks = temp;
	vnMinorTicks = temp;


	// Show all axes. 0 or 1, 1 for show.
	vnAxes[nAxisType] = bAxisOn ? 1:0;
	// Show all labels. 0 or 1, 1 for show.
	vnLabels[nAxisType] = bLabels ? 1:0;
	// Show axis titles of left and bottom axes. 0 or 1, 1 for show.
	vnTitles[nAxisType] = bTitleOn ? 1:0;
	// Set the major tick and minor tick of all axes as IN format
	// See other TICK_* items in graph_utils.h.
	// Options are:
	//TICK_IN_OUT,
	//TICK_IN,
	//TICK_OUT,
	//TICK_NONE
	vnMajorTicks[nAxisType] = nMajorTicks;
	vnMinorTicks[nAxisType] = nMinorTicks;
	gl_smart_show_object(gl, vnAxes, vnLabels, vnTitles, vnMajorTicks, vnMinorTicks);
}
*/
void DataPlotter::axis_from(int nAxisType, double dFrom)
{
	Tree tr;
	Axis axis = get_axis(nAxisType);
	tr.Root.Scale.From.dVal = dFrom;
	upate_axis(tr, axis);
}
void DataPlotter::axis_to(int nAxisType, double dTo)
{
	Tree tr;
	Axis axis = get_axis(nAxisType);
	tr.Root.Scale.To.dVal = dTo;
	upate_axis(tr, axis);
}


void DataPlotter::axis_increment_by_value(int nAxisType, double dIncrementBy)
{
	Tree tr;
	//tr = gl.YAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	
	Axis axis = get_axis(nAxisType);	
	tr.Root.Scale.Value.dVal = dIncrementBy;
	tr.Root.Scale.IncrementBy.nVal = 0; //increment by count
	upate_axis(tr, axis);
}
void DataPlotter::axis_increment_by_ticks(int nAxisType, int nMajTicksCount)
{
	
	Tree tr;
	//tr = gl.YAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	
	Axis axis = get_axis(nAxisType);
	tr.Root.Scale.MajorTicksCount.nVal = nMajTicksCount;

	tr.Root.Scale.IncrementBy.nVal = 1; //increment by count
	upate_axis(tr, axis);
	
}
void DataPlotter::smart_axis_increment(int nAxisType)
{
	
	Tree tr;
	//tr = gl.YAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	int nMajTicksCount = 5;
	Axis axis = get_axis(nAxisType);
	tr = axis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	double dFrom = tr.Root.Scale.From.dVal;
	double dTo = tr.Root.Scale.To.dVal;
	double diff = dTo-dFrom;
	double exp = log10(diff/nMajTicksCount);
	int iexp = round(exp,0);
	
	double dIncrementBy = pow(10,iexp);
	
	if(diff/dIncrementBy <= 3)
		dIncrementBy = dIncrementBy/2;
	else if(diff/dIncrementBy >= 9)
		dIncrementBy = dIncrementBy*2;
	axis_increment_by_value(nAxisType,dIncrementBy);
	
	//automatically determine the decimal places.
	int nDigits = abs(floor(log10(dTo))-floor(log10(dIncrementBy)));
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		if(tr.Root.Labels.BottomLabels.NumericFormat.nVal == 1)
		{
			tr.Root.Labels.BottomLabels.CustomDecimal.nVal = 1;
			tr.Root.Labels.BottomLabels.DecimalPlaces.nVal = nDigits;
		}
		break;
	case AXIS_LEFT:
		if(tr.Root.Labels.LeftLabels.NumericFormat.nVal == 1)
		{
			tr.Root.Labels.LeftLabels.CustomDecimal.nVal = 1;
			tr.Root.Labels.LeftLabels.DecimalPlaces.nVal = nDigits;
		}
		break;
	case AXIS_TOP:
		if(tr.Root.Labels.TopLabels.NumericFormat.nVal == 1)
		{
			tr.Root.Labels.TopLabels.CustomDecimal.nVal = 1;
			tr.Root.Labels.TopLabels.DecimalPlaces.nVal = nDigits;
		}
		break;
	case AXIS_RIGHT:
		if(tr.Root.Labels.RightLabels.NumericFormat.nVal ==1)
		{
			tr.Root.Labels.RightLabels.CustomDecimal.nVal = 1;
			tr.Root.Labels.RightLabels.DecimalPlaces.nVal = nDigits;
		}
			break;
	}
	upate_axis(tr, axis);
	
	

}


void DataPlotter::axis_rescale_type(int nAxisType, int nType)
{
	//Type
	Tree tr;
	Axis axis = get_axis(nAxisType);
	tr.Root.Scale.Type.nVal = nType;	
	upate_axis(tr, axis);
}
void DataPlotter::axis_rescale(int nAxisType, int nRescale)
{
	//Rescale
	Tree tr;
	Axis axis = get_axis(nAxisType);
	tr.Root.Scale.Rescale.nVal = nRescale;	
	upate_axis(tr, axis);
}
void DataPlotter::axis_rescale_margin(int nAxisType, double dResMargin)
{
	//rescale margin percentage
	Tree tr;
	Axis axis = get_axis(nAxisType);
	tr.Root.Scale.RescaleMargin.dVal = dResMargin;	
	upate_axis(tr, axis);
}


void DataPlotter::axis_label_size(int nAxisType, double dSize)
{
	//nAxisType: AXIS_LEFT, AXIS_RIGHT, AXIS_TOP, AXIS_BOTTOM
	//AxisObject ao = gl.YAxis.AxisObjects(AXISOBJPOS_AXIS_FIRST);
	Tree tr;
	//tr = gl.YAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//tr = gl.YAxis.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr,true,true);
	//tr.Root.Labels.LeftTitle.Text.strVal = "\b(%(?Y))";
	//tr.Root.Labels.LeftLabels.Font.Size.dVal = dSize;
	//tr.Root.Labels.LeftLabels.Font.Size.dVal = dSize;
	Axis axis = get_axis(nAxisType);
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		tr.Root.Labels.BottomLabels.Font.Size.dVal = dSize;
		break;
	case AXIS_LEFT:
		tr.Root.Labels.LeftLabels.Font.Size.dVal = dSize;
		break;
	case AXIS_TOP:
		tr.Root.Labels.TopLabels.Font.Size.dVal = dSize;
		break;
	case AXIS_RIGHT:
		tr.Root.Labels.RightLabels.Font.Size.dVal = dSize;
		break;
	}
	upate_axis(tr, axis);
	
}
void DataPlotter::axis_label_numeric_format(int nAxisType, int nFormat)
{
	
	
	Tree tr;
	Axis axis = get_axis(nAxisType);
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		tr.Root.Labels.BottomLabels.NumericFormat.nVal = nFormat;
		break;
	case AXIS_LEFT:
		tr.Root.Labels.LeftLabels.NumericFormat.nVal = nFormat;
		break;
	case AXIS_TOP:
		tr.Root.Labels.TopLabels.NumericFormat.nVal = nFormat;
		break;
	case AXIS_RIGHT:
		tr.Root.Labels.RightLabels.NumericFormat.nVal = nFormat;
		tr.Root.Labels.RightLabels.DecimalPlaces.nVal = 3;
		break;
	}
	upate_axis(tr, axis);
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	
}
void DataPlotter::axis_title_size(int nAxisType, double dSize)
{
	Tree tr;
	Axis axis = get_axis(nAxisType);
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		tr.Root.Titles.BottomTitle.Font.Size.dVal = dSize;
		break;
	case AXIS_LEFT:
		tr.Root.Titles.LeftTitle.Font.Size.dVal = dSize;
		break;
	case AXIS_TOP:
		tr.Root.Titles.TopTitle.Font.Size.dVal = dSize;
		break;
	case AXIS_RIGHT:
		tr.Root.Titles.RightTitle.Font.Size.dVal = dSize;
		break;
	}
	upate_axis(tr, axis);
}
void DataPlotter::axis_title_text(int nAxisType, string strText)
{
	
	Tree tr;
	Axis axis = get_axis(nAxisType);
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		tr.Root.Titles.BottomTitle.AddNode("Text").strVal = strText;////Text is an embedded data member of TreeNode class, but here Text is a subnode, use AddNode to specify it's a node but not a data member
		break;
	case AXIS_LEFT:
		tr.Root.Titles.LeftTitle.AddNode("Text").strVal = strText;
		break;
	case AXIS_TOP:
		tr.Root.Titles.TopTitle.AddNode("Text").strVal = strText;
		break;
	case AXIS_RIGHT:
		tr.Root.Titles.RightTitle.AddNode("Text").strVal = strText;
		break;
	}
	upate_axis(tr, axis);
}
void DataPlotter::axis_color(int nAxisType,DWORD dwColor)
{
	Tree tr;
		Axis axis = get_axis(nAxisType);
		switch(nAxisType)
		{
		case AXIS_BOTTOM:
			tr.Root.Ticks.BottomTicks.Color.nVal = dwColor;
			tr.Root.Labels.BottomLabels.Color.nVal = dwColor;
			tr.Root.Titles.BottomTitle.Color.nVal = dwColor;
			break;
		case AXIS_LEFT:
			tr.Root.Ticks.LeftTicks.Color.nVal = dwColor;
			tr.Root.Labels.LeftLabels.Color.nVal = dwColor;
			tr.Root.Titles.LeftTitle.Color.nVal = dwColor;
			break;
		case AXIS_TOP:
			tr.Root.Ticks.TopTicks.Color.nVal = dwColor;
			tr.Root.Labels.TopLabels.Color.nVal = dwColor;
			tr.Root.Titles.TopTitle.Color.nVal = dwColor;
			break;
		case AXIS_RIGHT:
			tr.Root.Ticks.RightTicks.Color.nVal = dwColor;
			tr.Root.Labels.RightLabels.Color.nVal = dwColor;
			tr.Root.Titles.RightTitle.Color.nVal = dwColor;
			break;
		}
		upate_axis(tr, axis);
}
void DataPlotter::axis_color_automatic(int nAxisType)
{
	Tree tr;
	Axis axis = get_axis(nAxisType);
	switch(nAxisType)
	{
	case AXIS_BOTTOM:
		tr.Root.Ticks.BottomTicks.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Labels.BottomLabels.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Titles.BottomTitle.Color.nVal = INDEX_COLOR_AUTOMATIC;
		break;
	case AXIS_LEFT:
		tr.Root.Ticks.LeftTicks.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Labels.LeftLabels.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Titles.LeftTitle.Color.nVal = INDEX_COLOR_AUTOMATIC;
		break;
	case AXIS_TOP:
		tr.Root.Ticks.TopTicks.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Labels.TopLabels.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Titles.TopTitle.Color.nVal = INDEX_COLOR_AUTOMATIC;
		break;
	case AXIS_RIGHT:
		tr.Root.Ticks.RightTicks.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Labels.RightLabels.Color.nVal = INDEX_COLOR_AUTOMATIC;
		tr.Root.Titles.RightTitle.Color.nVal = INDEX_COLOR_AUTOMATIC;
		break;
	}
	upate_axis(tr, axis);
}

void DataPlotter::axis_pos_offset(int nAxisType, double dPosOffset)
{
	AxisObject ao;
	switch (nAxisType)
	{
	case AXIS_BOTTOM:
		ao = gl.XAxis.AxisObjects(AXISOBJPOS_AXIS_FIRST);
		ao.SetPosition(AXIS_POS_REAL,ao.GetPosition()-dPosOffset);
		break;
	case AXIS_LEFT:
		ao = gl.YAxis.AxisObjects(AXISOBJPOS_AXIS_FIRST);
		ao.SetPosition(AXIS_POS_REAL,ao.GetPosition()-dPosOffset);
		break;
	case AXIS_TOP:
		ao = gl.XAxis.AxisObjects(AXISOBJPOS_AXIS_SECOND);
		ao.SetPosition(AXIS_POS_REAL,ao.GetPosition()+dPosOffset);
		break;
	case AXIS_RIGHT:
		
		ao = gl.YAxis.AxisObjects(AXISOBJPOS_AXIS_SECOND);
		ao.SetPosition(AXIS_POS_REAL,ao.GetPosition()+dPosOffset);
		break;
	}
	
}
void DataPlotter::make_plot(string sWksName, int nCx, int nCy, int nPlotID)
{
	//
	//DataRange is for getting and putting data from and to a worksheet, matrix, and graph window
	
	
	//nPlotID: 
	//IDM_PLOT_LINE for line;
	//IDM_PLOT_SCATTER for scatter;
	//IDM_PLOT_LINESYMB for line + symbol.
	
	//marker: "None", "o"
	//linestyle: "None", "-"
	//color: "b".It is an RGB color
	
	//Determine plot type
	Worksheet wks(sWksName);
	
	//Create data range
	DataRange dr;
	dr.Add(wks,nCx,"X");
	dr.Add(wks,nCy,"Y");
	
	
	int nPlot = gl.AddPlot(dr,nPlotID);
	
	dp = gl.DataPlots(nPlot);
	gl.Rescale();
	
	//dp.SetColor(RGB2OCOLOR(color));
	////Show tree of the dataplot
	//Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	/*
	Tree tr;
	tr.Root.Symbol.Size.dVal = markersize;
	tr.Root.Line.Width.dVal = linewidth;
	tr.Root.Symbol.Shape.nVal = marker;
	tr.Root.Line.Style.nVal = linestyle;
	int nErr = dp.UpdateThemeIDs(tr.Root);// allows a tree with no theme ID to be used in ApplyFormat
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
	gl.Rescale();
	*/
}

void DataPlotter::plot_marker_style(int nMarkerStyle)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	tr.Root.Symbol.Shape.nVal = nMarkerStyle;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_marker_size(double dMarkerSize)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	tr.Root.Symbol.Size.dVal = dMarkerSize;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_marker_edge_color(DWORD dwColor)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	tr.Root.Symbol.Interior.nVal = 1;//interior type
	tr.Root.Symbol.EdgeColor.nVal = dwColor;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_marker_edge_width(double dWidth)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	tr.Root.Symbol.Interior.nVal = 1;//interior type
	tr.Root.Symbol.EdgeWidth.dVal = dWidth;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_marker_face_color(DWORD dwColor)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	tr.Root.Symbol.Interior.nVal = 1;//interior type
	tr.Root.Symbol.FillColor.nVal = dwColor;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_line_color(DWORD dwColor)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	tr.Root.Line.Color.nVal = dwColor;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_line_style(int nLineStyle)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	tr.Root.Line.Style.nVal = nLineStyle;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}
void DataPlotter::plot_line_width(double dLineWidth)
{
	Tree tr;
	//tr = dp.GetFormat(FPB_ALL, FOB_ALL, true, true);
	//out_tree(tr);
	tr.Root.Line.Width.dVal = dLineWidth;
	int nErr = dp.UpdateThemeIDs(tr.Root);
	if(nErr ==0)
		dp.ApplyFormat(tr,true,true);
}

void DataPlotter::make_worksheetpage(string strWksPgName)
{
	wksPg.Create();
	wksPg.SetName(strWksPgName);
}
void DataPlotter::add_worksheet_from_csv(string strFileName, string strWksName)
{
    if(wksPg)
    {
		string fileName = strFileName;  // csv file browser
		if(!fileName.IsFile())
			return;
		
		// create a new worksheet
		int index = wksPg.AddLayer(strWksName);
		Worksheet wks = wksPg.Layers(index);
		// ascii import
		ASCIMP ai;
		initASCIMP(ai);  // initialize
		ai.iAutoSubHeaderLines = 1;  // auto detect subheader line
		ai.iDelimited = 1;  // use delimiter
		ai.iDelimiter = ASCIMP_DELIM_COMMA;	// comma as delimiter
		ai.iNonnumeric = 1; // NONNUMERIC_READ_AS_MISSING
		// special quote symbol and remove it when import
		ai.cQuote = '\"';
		ai.flags |= AI_FLAG_REMOVE_QUOTES;
		ai.iRenameWks = 0;
		int nret = AscImpReadFileStruct(fileName, &ai, AIRF_USE_ASCIMP);
		if(0 == nret)
		{
			// import
			wks.ImportASCII(fileName, ai);
		}   
    }
}
