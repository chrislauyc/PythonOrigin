enum //Axis_Type
{
	SELECT_Y_AXIS,
	SELECT_X_AXIS
};
enum //INCREMENT_TYPE
{
	INCREMENT_BY_VALUE,
	INCREMENT_BY_TICK	
};
/*Colors
INDEX_COLOR_AUTOMATIC
SYSCOLOR_BLACK
SYSCOLOR_RED
SYSCOLOR_GREEN
SYSCOLOR_BLUE
SYSCOLOR_CYAN
SYSCOLOR_MAGENTA 
SYSCOLOR_YELLOW
SYSCOLOR_DKYELLOW
SYSCOLOR_NAVY
SYSCOLOR_PURPLE
SYSCOLOR_WINE
SYSCOLOR_OLIVE
SYSCOLOR_DKCYAN
SYSCOLOR_ROYAL
SYSCOLOR_ORANGE
SYSCOLOR_VIOLET
SYSCOLOR_PINK
SYSCOLOR_WHITE
SYSCOLOR_LTGRAY
SYSCOLOR_GRAY
SYSCOLOR_LTYELLOW
SYSCOLOR_LTCYAN
SYSCOLOR_LTMAGENTA
SYSCOLOR_DKGRAY
*/
class DataPlotter
{
private:
	//get axis options: AXIS_BOTTOM, AXIS_LEFT, AXIS_TOP, AXIS_RIGHT
	Axis get_axis(int nAxisType);
	//use the format tree to update format
	void update_axis(Tree tr, Axis axis);
public:
	//##############Member data########################
	GraphPage gp; //Holds the current graph page object
	GraphLayer gl; //Holds the current graph layer object
	DataPlot dp; //Holds the current dataplot object
	WorksheetPage wksPg;//Holds the current worksheet page
	bool x_linked; //Important! Currently, it only supports either linking all of the layers or not.
	bool y_linked;
	//#############Member Functions#####################
	//Constructor takes in a new name for graph page
	DataPlotter(string sGraphPageName);
	//Use layer index to set the current graph layer contained in gl
	void select_layer(int nLayerInd);
	//Use the dataplot index to set the current dataplot contained in dp
	void select_plot(int nPlotInd);
	//add layer
	void add_layer(int nAxisType, int nLinkTo, int nXAxisLink, int nYAxisLink);
	
	//****************Layer Settings*******************
	void graphpage_resize(double dWidth, double dHeight);//This doesn't work
	//****************Axis Settings********************
	void show_axis(int nAxisType, bool bAxisOn, bool bLabels, bool bTitleOn, int nMajorTicks, int nMinorTicks);
	//set axis range
	void axis_from(int nAxisType, double dFrom);
	void axis_to(int nAxisType, double dTo);
	//set increment
	void axis_increment_by_value(int nAxisType, double dIncrementBy); //increment by value
	void axis_increment_by_ticks(int nAxisType, int nMajTicksCount); //increment by ticks
	void smart_axis_increment(int nAxisType); //increment by ticks
	//set rescale parameters
	void axis_rescale_type(int nAxisType, int nType); //Linear: 0, see origin lab
	void axis_rescale(int nAxisType, int nRescale); //Fixed: 0, see origin lab
	void axis_rescale_margin(int nAxisType, double dResMargin);//default: 8%
	
	//set axis label
	void axis_label_size(int nAxisType, double dSize);
	//set axis label Numeric Format
	//nFormat: 0=Decimal(1000),1:Scientific(1E^3),2:Engineering(3k),3:Decimal(1,000),4:Scientific(1E3)
	void axis_label_numeric_format(int nAxisType, int nFormat);
	//set axis title
	void axis_title_size(int nAxisType, double dSize);
	void axis_title_text(int nAxisType, string strText);
	//set axis color
	void axis_color(int nAxisType,DWORD dwColor);
	//set axis color same as data
	void axis_color_automatic(int nAxisType);
	//set axis position offset
	void axis_pos_offset(int nAxisType, double dPosOffset); //nAxisType: AXIS_TOP, AXIS_BOTTOM, AXIS_LEFT, AXIES_RIGHT
	
	//***************DataPlot Settings********************
	//plot (would include color, size, and plot type) (specific worksheet name)
	void make_plot(string sWksName, int nCx, int nCy, int nPlotID); //IDM_PLOT_LINE, IDM_PLOT_SCATTER, IDM_PLOT_LINESYMB
	//
	void plot_marker_style(int nMarkerStyle);
	void plot_marker_size(double dMarkerSize);
	void plot_marker_edge_color(DWORD dwColor);
	void plot_marker_edge_width(double dWidth);
	void plot_marker_face_color(DWORD dwColor);
	
	void plot_line_color(DWORD dwColor);
	void plot_line_style(int nLineStyle);
	void plot_line_width(double dLineWidth);
	//****************Worksheet*************************
	void make_worksheetpage(string strWksPgName);
	void add_worksheet_from_csv(string strFileName, string strWksName);
	
	
};