/*------------------------------------------------------------------------------*
 * File Name: Plotter															*
 * Creation: Chris Lau															*
 * Purpose: OriginC Source C++ file. Plot NPMS figures							*
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
//#include <..\Originlab\graph_utils.h> 
#include <E:\OneDrive - University of Utah\Anderson's lab\OriginC\graph_utils.h> 

////////////////////////////////////////////////////////////////////////////////////

//#pragma labtalk(0) // to disable OC functions for LT calling.

////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.

#include "DataPlotter.h"
////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

AxisObject get_ao(const GraphLayer gl, const int nAxisType)
{
	AxisObject ao;
	switch (nAxisType)
	{
	case AXIS_BOTTOM:
		ao = gl.XAxis.AxisObjects(AXISOBJPOS_AXIS_FIRST);
		break;
	case AXIS_LEFT:
		ao = gl.YAxis.AxisObjects(AXISOBJPOS_AXIS_FIRST);
		break;
	case AXIS_TOP:
		ao = gl.XAxis.AxisObjects(AXISOBJPOS_AXIS_SECOND);
		break;
	case AXIS_RIGHT:
		ao = gl.YAxis.AxisObjects(AXISOBJPOS_AXIS_SECOND);
		break;
	}
	return ao;
}
void clear_project()
{
	//empty the project space
	bool exists = true;
	WorksheetPage wksPg;
	do
	{
		wksPg = Project.WorksheetPages.Item(0);
		
		if(wksPg)
			wksPg.Destroy();
		else
			exists = false;
	}while(exists);
	exists = true;
	GraphPage grPg;
	do
	{
		grPg = Project.GraphPages.Item(0);
		if(grPg)
			grPg.Destroy();
		else
			exists = false;
	}while(exists);
}

WorksheetPage make_wb(string name,string longname = "")
{
	WorksheetPage wksPg;
	wksPg.Create();
	wksPg.SetName(name);
	wksPg.SetLongName(longname,false);
	return wksPg;
}
Worksheet Add_ws(WorksheetPage wksPg, string name)
{
	int index = wksPg.AddLayer(name);
	Worksheet wksNew = wksPg.Layers(index);
	return wksNew;
}



GraphPage make_gp(string name)
{	
	//numYs is the number of y axies to add

	GraphPage gp(name);
	//need to be careful about this. May need to remove.
	if(gp)
		gp.Destroy();
	//
	gp.Create();
	gp.SetName(name);
	
	return gp;
}
//Newly added





// Newly added END

GraphLayer add_layer(GraphPage gp, string sAxType = "Right", int nLinkTo = 0, int nXAxisLink = LINK_STRAIGHT, int nYAxisLink = 0)
{
	// sAxType will determine which axis will be visible. Options are "Right","Left","Top","Bottom"
	// int nLinkTo is the index of the layer to link to
	// int nXAxisLink is how the x axis should be linked. Options are: 0, LINK_STRAIGHT, LINKED_AXIS_CUSTOM, LINKED_AXIS_ALIGN
	// int nYAxisLink is how the y axis should be linked. Options are: 0, LINK_STRAIGHT, LINKED_AXIS_CUSTOM, LINKED_AXIS_ALIGN
	bool bBottom = false, bLeft = false, bTop = false, bRight = false;
	int nType;
	if(sAxType == "Left")
	{
		bLeft = true;
		nType = AXIS_LEFT;
	}
	else if(sAxType == "Right")
	{
		bRight = true;
		nType = AXIS_RIGHT;
	}
	else if(sAxType == "Bottom")
	{
		bBottom = true;
		nType = AXIS_BOTTOM;
	}
	else if(sAxType == "Top")
	{
		bTop = true;
		nType = AXIS_TOP;
	}
	else
	{
		bRight = true;
		nType = AXIS_RIGHT;
	}
	
	
	//int nLinkTo = 0; // 
	bool bActivateNewLayer = false;
	page_add_layer(gp, bBottom, bLeft, bTop, bRight,ADD_LAYER_INIT_SIZE_POS_MOVE_OFFSET, bActivateNewLayer, nLinkTo, nXAxisLink, nYAxisLink);
	
	GraphLayer gl = gp.Layers(gp.Layers.Count()-1);
	
	AxisObject ao = get_ao(gl,nType);
	
	
	return gl;
}


/*
int make_plot(GraphLayer gl,DataRange dr, string marker = "o", string linestyle = "-", DWORD color = 0)
{
	//
	//DataRange is for getting and putting data from and to a worksheet, matrix, and graph window
	
	
	//nPlotID: 
	//IDM_PLOT_LINE for line;
	//IDM_PLOT_SCATTER for scatter;
	//IDM_PLOT_LINESYMB for line + symbol.
	
	//marker: "None", "o"
	//linestyle: "None", "-"
	//color: "b"
	
	//Determine plot type
	int nPlotID = IDM_PLOT_LINESYMB;
	if(marker != "None" && linestyle != "None")
		nPlotID = IDM_PLOT_LINESYMB;
	else if(marker == "None" && linestyle != "None")
		nPlotID = IDM_PLOT_LINE;
	else if(marker != "None" && linestyle == "None")
		nPlotID = IDM_PLOT_SCATTER;
	
	int nPlot = gl.AddPlot(dr,nPlotID);
	//Convert RGB color in ocolor
	DataPlot dp = gl.DataPlots(nPlot);
	dp.SetColor(RGB2OCOLOR(color));
	
	
	return 0;
}
int make_dr(Worksheet wks, int nCx=0, int nCy=0)
{
	DataRange dr;
	dr.Add(wks,nCx,"X");
	dr.Add(wks,nCy,"Y");
	return dr;
}
*/

int Plotter()
{
	//Important!
	//make sure to run run.loadoc(Originlab\graph_utils.c, 16); in labtalk for page_add_layer to work
	//run run.addoc(E:\OneDrive - University of Utah\Anderson's lab\OriginC\Plotter.cpp); in labtalk to import this file
	
	//abs path: E:\OneDrive - University of Utah\Anderson's lab\OriginC\Plotter.cpp
	//clear_project();
	//WorksheetPage wbsPg = make_wb("Experiment");
	//Worksheet msheet = Add_ws(wbsPg, "Mass");
	//Worksheet tsheet = Add_ws(wbsPg, "Temperature");
	
	
	//GraphPage gp = make_gp("Summary");
	//GraphLayer g1 = add_layer(gp);
	//GraphLayer g2 = add_layer(gp,"Left",1);
	
	
	//AxisObject ao1 = get_ao(g1,AXIS_RIGHT);
	//AxisObject ao2 = get_ao(g2,AXIS_RIGHT);
	//double pos1 = ao1.GetPosition();
	//ao2.SetPosition(AXIS_POS_REAL, pos1+2);
	
	//GraphLayer gl = gp.Layers(0);
	//DWORD color = RGB(0,0,0);
	//DataRange dr;
	//make_plot(gl,dr,"o","-",color);
	//Worksheet msheet("[Experiment]Mass");
	//Worksheet tsheet("[Experiment]Temp");
	
	//gp.Layers.Count()
	
	DataPlotter dp1("try1");
	//dp1.add_layer("Right",0,LINK_STRAIGHT,0);
	//dp1.select_layer(2);
	
	dp1.make_plot("[Experiment]Mass",0,11,IDM_PLOT_LINESYMB);
	dp1.show_axis(AXIS_LEFT, true, true, true, TICK_OUT, TICK_OUT);
	dp1.axis_label_size(AXIS_LEFT,15);
	dp1.axis_title_size(AXIS_LEFT,20);
	//dp1.axis_title_text(AXIS_LEFT,"%(?Y)");
	dp1.axis_label_numeric_format(AXIS_LEFT,1);
	dp1.axis_increment_by_ticks(AXIS_LEFT,7);
	dp1.plot_line_color(SYSCOLOR_RED);
	dp1.plot_marker_edge_color(SYSCOLOR_RED);
	dp1.axis_color_automatic(AXIS_LEFT);
	
	return 0;
}



