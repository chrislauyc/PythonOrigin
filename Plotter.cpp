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

////////////////////////////////////////////////////////////////////////////////////

//#pragma labtalk(0) // to disable OC functions for LT calling.


////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.
#include <..\Originlab\graph_utils.h>
#include <..\originlab\fu_utils.h>
#include "DataPlotter.h"
////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.


//need these libraries at install
//run.LoadOC(Originlab\\fu_utils.c);
//run.loadoc(Originlab\\graph_utils.c, 16);


DataPlotter dataplotter;

void make_worksheetpage(string strWksPgName)
{
	dataplotter.make_worksheetpage(strWksPgName);
}
void add_worksheet_from_csv(string strFileName, string strWksName)
{
	dataplotter.add_worksheet_from_csv(strFileName, strWksName);
}
void make_graph(string sGraphPageName)
{
	DataPlotter new_dp(sGraphPageName);
	dataplotter = new_dp;
	dataplotter.show_axis(AXIS_BOTTOM, true, true, true, TICK_OUT, TICK_OUT);
	dataplotter.show_axis(AXIS_LEFT, true, true, true, TICK_OUT, TICK_OUT);
}
void graphpage_resize(double dWidth, double dHeight)
{
	//This doesn't work. Have to fit graph to layer manually.
	//dataplotter.graphpage_resize(dWidth,dHeight);
}
void add_xlinked_layer_right()
{
	dataplotter.add_layer(AXIS_RIGHT,0,LINK_STRAIGHT,0);
	dataplotter.show_axis(AXIS_RIGHT, true, true, true, TICK_OUT, TICK_OUT);
}
void add_xlinked_layer_left()
{
	dataplotter.add_layer(AXIS_LEFT,0,LINK_STRAIGHT,0);
	dataplotter.show_axis(AXIS_LEFT, true, true, true, TICK_OUT, TICK_OUT);
}
void xrange(double dFrom, double dTo)
{
	dataplotter.axis_from(AXIS_BOTTOM,dFrom);
	dataplotter.axis_to(AXIS_BOTTOM,dTo);
	dataplotter.smart_axis_increment(AXIS_BOTTOM);
}
void yrange(double dFrom, double dTo)
{
	dataplotter.axis_from(AXIS_RIGHT,dFrom);
	dataplotter.axis_from(AXIS_LEFT,dFrom);
	dataplotter.axis_to(AXIS_RIGHT,dTo);
	dataplotter.axis_to(AXIS_LEFT,dTo);
	dataplotter.smart_axis_increment(AXIS_RIGHT);
	dataplotter.smart_axis_increment(AXIS_LEFT);
}

void xaxis_label_size(double dSize)
{
	dataplotter.axis_label_size(AXIS_BOTTOM,dSize);
}
void yaxis_label_size(double dSize)
{
	dataplotter.axis_label_size(AXIS_LEFT,dSize);
	dataplotter.axis_label_size(AXIS_RIGHT,dSize);
}
void xaxis_label_numeric_format(int nFormat)
{
	dataplotter.axis_label_numeric_format(AXIS_BOTTOM, nFormat);
	dataplotter.smart_axis_increment(AXIS_BOTTOM);
}
void yaxis_label_numeric_format(int nFormat)
{
	dataplotter.axis_label_numeric_format(AXIS_LEFT, nFormat);
	dataplotter.axis_label_numeric_format(AXIS_RIGHT, nFormat);
	dataplotter.smart_axis_increment(AXIS_LEFT);
	dataplotter.smart_axis_increment(AXIS_RIGHT);
}
void xtitle(string strText)
{
	dataplotter.axis_title_text(AXIS_BOTTOM, strText);
}
void ytitle(string strText)
{
	dataplotter.axis_title_text(AXIS_LEFT, strText);
	dataplotter.axis_title_text(AXIS_RIGHT, strText);
}
void xtitle_size(double dSize)
{
	dataplotter.axis_title_size(AXIS_BOTTOM,dSize);
}
void ytitle_size(double dSize)
{
	dataplotter.axis_title_size(AXIS_LEFT, dSize);
	dataplotter.axis_title_size(AXIS_RIGHT, dSize);
}
void xaxis_color(int nR, int nG, int nB)
{
	DWORD dwColor = RGB2OCOLOR(RGB(nR,nG,nB));
	dataplotter.axis_color(AXIS_BOTTOM,dwColor);
	dataplotter.axis_color(AXIS_TOP,dwColor);
}
void yaxis_color(int nR, int nG, int nB)
{
	DWORD dwColor = RGB2OCOLOR(RGB(nR,nG,nB));
	dataplotter.axis_color(AXIS_LEFT,dwColor);
	dataplotter.axis_color(AXIS_RIGHT,dwColor);
}
void yaxis_color_automatic()
{
	dataplotter.axis_color_automatic(AXIS_LEFT);
	dataplotter.axis_color_automatic(AXIS_RIGHT);
}
void yaxis_pos_offset_left(double dPosOffset)
{
	dataplotter.axis_pos_offset(AXIS_LEFT, dPosOffset);
}
void yaxis_pos_offset_right(double dPosOffset)
{
	dataplotter.axis_pos_offset(AXIS_RIGHT, dPosOffset);
}
void make_linesymb_plot(string sWksName, int nCx, int nCy)
{
	dataplotter.make_plot(sWksName,nCx,nCy,IDM_PLOT_LINESYMB);
	dataplotter.smart_axis_increment(AXIS_BOTTOM);
	dataplotter.smart_axis_increment(AXIS_LEFT);
}
void make_line_plot(string sWksName, int nCx, int nCy)
{
	dataplotter.make_plot(sWksName,nCx,nCy,IDM_PLOT_LINE);
	dataplotter.smart_axis_increment(AXIS_BOTTOM);
	dataplotter.smart_axis_increment(AXIS_LEFT);
}
void make_scatter_plot(string sWksName, int nCx, int nCy)
{
	dataplotter.make_plot(sWksName,nCx,nCy,IDM_PLOT_SCATTER);
	dataplotter.smart_axis_increment(AXIS_BOTTOM);
	dataplotter.smart_axis_increment(AXIS_LEFT);
}

void plot_marker_style(int nMarkerStyle)
{
	dataplotter.plot_marker_style(nMarkerStyle);
}
void plot_marker_size(double dMarkerSize)
{
	dataplotter.plot_marker_size(dMarkerSize);
}
void plot_marker_edge_color(int nR,int nG, int nB)
{
	DWORD dwColor = RGB2OCOLOR(RGB(nR,nG,nB));
	dataplotter.plot_marker_edge_color(dwColor);
}
void plot_marker_edge_width(double dWidth)
{
	dataplotter.plot_marker_edge_width(dWidth);
}
void plot_marker_face_color(int nR,int nG, int nB)
{
	DWORD dwColor = RGB2OCOLOR(RGB(nR,nG,nB));
	dataplotter.plot_marker_face_color(dwColor);
}
void plot_line_color(int nR,int nG, int nB)
{
	DWORD dwColor = RGB2OCOLOR(RGB(nR,nG,nB));
	dataplotter.plot_line_color(dwColor);
}
void plot_line_style(int nLineStyle)
{
	dataplotter.plot_line_style(nLineStyle);
}
void plot_line_width(double dLineWidth)
{
	dataplotter.plot_line_width(dLineWidth);
}
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
	
	/*
	DataPlotter dp1("try3");
	dp1.make_plot("[Experiment]Mass",0,11,IDM_PLOT_LINESYMB);
	dp1.axis_to(AXIS_LEFT,11400000);
	dp1.smart_axis_increment(AXIS_LEFT);
	*/
	
	
	/*
	DataPlotter dp1("try2");
	//dp1.add_layer("Right",0,LINK_STRAIGHT,0);
	//dp1.select_layer(2);
	
	dp1.make_plot("[Experiment]Mass",0,11,IDM_PLOT_LINESYMB);
	dp1.show_axis(AXIS_LEFT, true, true, true, TICK_OUT, TICK_OUT);
	dp1.show_axis(AXIS_BOTTOM, true, true, true, TICK_OUT, TICK_OUT);
	dp1.axis_label_size(AXIS_LEFT,15);
	dp1.axis_label_size(AXIS_BOTTOM,15);
	dp1.axis_title_size(AXIS_LEFT,20);
	dp1.axis_title_size(AXIS_BOTTOM,20);
	//dp1.axis_title_text(AXIS_LEFT,"%(?Y)");
	dp1.axis_label_numeric_format(AXIS_LEFT,1);
	dp1.axis_increment_by_ticks(AXIS_LEFT,7);
	dp1.plot_line_color(SYSCOLOR_BLUE);
	dp1.plot_marker_edge_color(SYSCOLOR_BLUE);
	dp1.axis_color_automatic(AXIS_LEFT);
	
	
	dp1.add_layer(AXIS_RIGHT, 0, LINK_STRAIGHT, 0);
	dp1.make_plot("[Experiment]Mass",0,9,IDM_PLOT_LINESYMB);
	dp1.plot_marker_edge_color(SYSCOLOR_RED);
	dp1.axis_color_automatic(AXIS_RIGHT);
	dp1.plot_marker_size(3);
	dp1.axis_pos_offset(AXIS_RIGHT, 3);
	*/
	return 0;
}



