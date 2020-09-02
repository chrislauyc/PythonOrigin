
import OriginExt
import time
import os
import numpy as np
import pandas as pd
import matplotlib.colors as colors
def connect_to_origin():
    # Connect to Origin client
    # OriginExt.Application() forces a new connection
    origin = OriginExt.ApplicationSI()
    origin.Visible = origin.MAINWND_SHOW # Make session visible
    # Session can be later closed using origin.Exit()
    # Close previous project and make a new one
    # origin.NewProject()
    # Wait for origin to compile
    # https://www.originlab.com/doc/LabTalk/ref/Second-cmd#-poc.3B_Pause_up_to_the_specified_number_of_seconds_to_wait_for_Origin_OC_startup_compiling_to_finish
    origin.Execute("sec -poc 3.5")
    time.sleep(3.5)
    folder_path = os.getcwd()
    code_path = os.path.join(folder_path,'Plotter.cpp')
    origin.Execute('run.loadoc({},16);'.format(code_path))
    time.sleep(3.5)
    
    return origin
def disconnect_from_origin(origin):
    origin.Exit()
def get_origin_version(origin):
    # Get origin version
    # Origin 2015 9.2xn
    # Origin 2016 9.3xn
    # Origin 2017 9.4xn
    # Origin 2018 >= 9.50n and < 9.55n
    # Origin 2018b >= 9.55n
    # Origin 2019 >= 9.60n and < 9.65n (Fall 2019)
    # Origin 2019b >= 9.65n (Spring 2020)
    return origin.GetLTVar("@V")
def save_project(origin,project_name,full_path):
    # File ending is automatically added by origin
    project_name = project_name.replace('.opju','').replace('.opj','')
    origin.Execute("save " + os.path.join(full_path,project_name))


def numpy_to_origin(
    data_array,column_axis=0,types=None,
    long_names=None,comments=None,units=None,
    user_defined=None,origin=None,
    worksheet_name='Sheet',workbook_name='Book'):
    '''
    Sends 2d numpy array to originlab worksheet
    Inputs:
    data_array = numpy array object
    column_axis = integer (0 or 1) for axis to interpret as worksheet columns
    long_names,comments,units = lists for header rows, length = # of columns
    user_defined = list of (key,value) tuples for metadata for a sheet
        e.g. [('Test Date','2019-01-01'),('Device Label','A12')]
    origin = origin session, which is returned from previous calls to this program
             if passed, a new session will not be created, and graph will be added to 
             current session
    origin_version = 2016 other year, right now >2016 handles DataRange differently
    types = column types, either 'x','y','x_err','y_err','z','label', or 'ignore'
    '''
    # If no origin session has been passed, start a new one
    if origin==None:
        origin = connect_to_origin()
    origin_version = get_origin_version(origin)
    # Check if workbook exists. If not create a new workbook page with this name
    layer_idx=None
    if origin.WorksheetPages(workbook_name) is None:
        workbook_name = origin.CreatePage(2, workbook_name , 'Origin') # 2 for workbook
        # Use Sheet1 if workbook is newly made
        layer_idx=0
    # get workbook instance from name
    wb = origin.WorksheetPages(workbook_name)
    if layer_idx is None:
        wb.Layers.Add() # Add a worksheet
        #then find the last worksheet to modify (to avoid overwriting other data)
        layer_idx = wb.Layers.Count - 1
    ws=wb.Layers(layer_idx) # Get worksheet instance, index starts at 0.
    ws.Name=worksheet_name # Set worksheet name
    # For now, assume only x and y data for each line (ignore error data)
    ws.Cols=data_array.shape[column_axis] # Set number of columns in worksheet
    # Change column Units, Long Name, or Comments]
    for col_idx in range(0,data_array.shape[column_axis]):
        col=ws.Columns(col_idx) # Get column instance, index starts at 0
        # Go through, check that each value exists and add to worksheet
        if (not long_names is None) and (len(long_names)>col_idx):
            col.LongName=long_names[col_idx]
        if (not units is None) and (len(units)>col_idx):
            col.Units=units[col_idx]
        if (not comments is None) and (len(comments)>col_idx):
            col.Comments=comments[col_idx]
        if not (types is None) and (len(types)>col_idx):
            type_str_to_int={'x':3,'y':0,'x_err':6,'y_err':2,'label':4,'z':5,'ignore':1}
            # Set column data type to (0 = Y, 1 = disregard, 2 = Y Error, 3 = X, 4 = Label, 5 = Z, and 6 = X Error.)
            # documentation here is off by one  https://www.originlab.com/doc/LabTalk/ref/Wks-Col-obj
            col.Type=type_str_to_int[types[col_idx].lower()]
        # Check dimensionality off array.
        # If one dimensional, each element is assumed to be a column
        # If two dimensional, check
        # other dimensions are not supported.
        if data_array.ndim == 2:
            if column_axis == 0:
                origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(data_array[col_idx,:]).tolist(), 0, col_idx) # start row, start col
            elif column_axis == 1:
                origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(data_array[:,col_idx]).tolist(), 0, col_idx) # start row, start col
        elif data_array.ndim == 1:
            origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(data_array[col_idx]).tolist(), 0, col_idx) # start row, start col
        else:
            print('only 1 and 2 dimensional arrays supported')
    #origin.PutWorksheet('['+wb.Name+']'+ws.Name, np.float64(data_array).T.tolist(), 0, col_idx) # start row, start col
    if not user_defined is None:
        # User Param Rows
        for idx,param in enumerate(user_defined):
            ws.Execute('wks.UserParam' + str(idx+1) + '=1; wks.UserParam' + str(idx+1) + '$="' + param[0] + '";')
            ws.Execute('col(1)[' + param[0] + ']$="' + param[1] + '";')
        origin.Execute('wks.col1.width=10;')
    return origin,wb,ws
###################################################################
#below are the python wrappers for functions written in originC
###################################################################

def load_plotter():
    pass
def make_graph(sGraphPageName,origin=None):
    origin.Execute('make_graph({});'.format(sGraphPageName))
def show_axis(nAxisType,origin=None):
    assert(nAxisType==0 or nAxisType==1 or nAxisType==2 or nAxisType==3)
    #nAxisType: 0=AXIS_BOTTOM, 1=AXIS_LEFT, 2=AXIS_TOP, 3=AXIES_RIGHT
    origin.Execute('show_axis({});'.format(nAxisType))
def select_graphpage(sGraphPageName,origin=None):
    origin.Execute('select_graphpage({});'.format(sGraphPageName))
def select_layer(nLayerInd,origin=None):
    origin.Execute('select_layer({});'.format(nLayerInd))
def select_plot(nPlotInd,origin=None):
    origin.Execute('select_plot({});'.format(nPlotInd))
def add_reflines_ver(lsReflines,origin=None):
    #input is a list of refline positions
    strReflines = ''
    for i,e in enumerate(lsReflines):
        if i > 0:
            strReflines += ' '
        strReflines += e
    if len(strReflines) != 0:
        origin.Execute('add_reflines_ver({});'.format(strReflines))
def add_reflines_hor(strReflines,origin=None):
    origin.Execute('add_reflines_hor({});'.format(strReflines))
def refline_fill_ver(nRefLineIndex,nFillToIndex,nR,nG,nB,origin=None):
    origin.Execute('refline_fill_ver({}{}{}{}{});'.format(nRefLineIndex,nFillToIndex,nR,nG,nB))    
def refline_fill_hor(nRefLineIndex,nFillToIndex,nR,nG,nB,origin=None):
    origin.Execute('refline_fill_hor{{}{}{}{}{});'.format(nRefLineIndex,nFillToIndex,nR,nG,nB))
def add_xlinked_layer_right(origin=None):
    origin.Execute('add_xlinked_layer_right();')
def add_xlinked_layer_left(origin=None):
    origin.Execute('add_xlinked_layer_left();')
def xrange(dFrom,dTo,origin=None):
    origin.Execute('xrange({},{});'.format(dFrom,dTo))
def yrange(dFrom,dTo,origin=None):
    origin.Execute('yrange({},{});'.format(dFrom,dTo))
def xsmart_axis_increment(origin=None):
    origin.Execute('xsmart_axis_increment();')
def ysmart_axis_increment(origin=None):
    origin.Execute('ysmart_axis_increment();')
def axis_rescale_type(nAxisType,nType,origin=None):
    origin.Execute('axis_rescale_type({},{});'.format(nAxisType,nType))
def axis_rescale(nAxisType,nRescale,origin=None):
    origin.Execute('axis_rescale({},{});'.format(nAxisType,nRescale))
def axis_rescale_margin(nAxisType,dResMargin,origin=None):
    origin.Execute('axis_rescale_margin({},{});'.format(nAxisType,dResMargin))
def xaxis_label_size(dSize,origin=None):
    origin.Execute('xaxis_label_size({});'.format(dSize))
def yaxis_label_size(dSize,origin=None):
    origin.Execute('yaxis_label_size({});'.format(dSize))
def xaxis_label_numeric_format(nFormat,origin=None):
    origin.Execute('xaxis_label_numeric_format({});'.format(nFormat))
def yaxis_label_numeric_format(nFormat,origin=None):
    origin.Execute('yaxis_label_numeric_format({});'.format(nFormat))
def xtitle(strText,origin=None):
    origin.Execute('xtitle({});'.format(strText))
def ytitle(strText,origin=None):
    origin.Execute('ytitle({});'.format(strText))
def xtitle_size(dSize,origin=None):
    origin.Execute('xtitle_size({});'.format(dSize))
def ytitle_size(dSize,origin=None):
    origin.Execute('ytitle_size({});'.format(dSize))
def xaxis_color(nR,nG,nB,origin=None):
    origin.Execute('xaxis_color({},{},{});'.format(nR,nG,nB))
def yaxis_color(nR,nG,nB,origin=None):
    origin.Execute('yaxis_color({},{},{});'.format(nR,nG,nB))
def yaxis_color_automatic(origin=None):
    origin.Execute('yaxis_color_automatic();')
def yaxis_pos_offset_left(dPosOffset,origin=None):
    origin.Execute('yaxis_pos_offset_left({});'.format(dPosOffset))
def yaxis_pos_offset_right(dPosOffset,origin=None):
    origin.Execute('yaxis_pos_offset_right({});'.format(dPosOffset))
def make_linesymb_plot(sWksName,nCx,nCy,origin=None):
    origin.Execute('make_linesymb_plot({},{},{});'.format(sWksName,nCx,nCy))
def make_line_plot(sWksName,nCx,nCy,origin=None):
    origin.Execute('make_line_plot({},{},{});'.format(sWksName,nCx,nCy))
def make_scatter_plot(sWksName,nCx,nCy,origin=None):
    origin.Execute('make_scatter_plot({},{},{});'.format(sWksName,nCx,nCy))
def plot_marker_style(nMarkerStyle,origin=None):
    origin.Execute('plot_marker_style({});'.format(nMarkerStyle))
def plot_marker_size(dMarkerSize,origin=None):
    origin.Execute('plot_marker_size({});'.format(dMarkerSize))
def plot_marker_edge_color(nR,nG,nB,origin=None):
    origin.Execute('plot_marker_edge_color({},{},{});'.format(nR,nG,nB))
def plot_marker_edge_width(dWidth,origin=None):
    origin.Execute('plot_marker_edge_width({});'.format(dWidth))
def plot_marker_face_color(nR,nG,nB,origin=None):
    origin.Execute('plot_marker_face_color({},{},{});'.format(nR,nG,nB))
def plot_line_color(nR,nG,nB,origin=None):
    origin.Execute('plot_line_color({},{},{});'.format(nR,nG,nB))
def plot_line_style(nR,nG,nB,origin=None):
    origin.Execute('plot_line_style({},{},{});'.format(nR,nG,nB))
def plot_line_width(dLineWidth,origin=None):
    origin.Execute('plot_line_width({});'.format(dLineWidth))
#############################################################
#Below is a higher level class for the user
#############################################################

class Py2OriginC():
    def __init__(self,origin):
        self.origin = origin
        self.GraphPages = []
    def connect(self):
        self.origin = connect_to_origin()
    def new_GraphPage(self,name):
        self.origin.CreatePage(3,name)
        gp = self.origin.GraphPages(name)
        gp = GraphPage(gp,self.origin)
        self.GraphPages.append(gp)
        return gp
    def fit_page_to_layers(self,gp):
        gp.Layers(0).Activate() #for some reason, this has to be done before calling the x-function
        self.origin.Execute('pfit2l margin:=tight;') #x-function to fit page to layers
    def new_WorkSheet(self,*args):
        return WorkSheet(*args,self.origin)
class WorkSheet(Py2OriginC):
    def __init__(self,ws_name,wb_name,origin):
        Py2OriginC.__init__(self,origin)
        self.ws = None
        if self.origin.WorksheetPages(wb_name) == None:
            self.origin.CreatePage(2,wb_name,'Origin')
        wb = self.origin.WorksheetPages(wb_name)
        for i in range(wb.Layers.Count):
            x = wb.Layers(i)
            if x.Name == ws_name:
                self.ws = x
        if self.ws == None:
            wb.Layers.Add()
            self.ws = wb.Layers(wb.Layers.Count-1)
            self.ws.Name = ws_name
    def destroy(self):
        self.ws.Destroy()
    def get_name(self):
        return self.ws.Name
    def from_df(self, df, units = None):
        if units != None:
            assert(len(units) != len(df.columns))
        self.ws.Cols = len(df.columns)
        for i, c in enumerate(df.columns):
            #df = pd.DataFrame() #this is for syntax checking
            #print(self.ws.Columns.__dict__)
            #self.ws.Columns.Add()
            col = self.ws.Columns(i)
            #col = self.ws.Columns(self.ws.Columns.Count-1)
            col.LongName = c
            col.SetData(np.float64(df[c])) #This needs some testing
            if units != None:
                col.Units = units[i]
            #col.Units = 'something'        
    name = property(get_name)
    
class GraphPage(Py2OriginC):
    def __init__(self,gp,origin=None):# pass in the actual object 
        Py2OriginC.__init__(self,origin)
        self.gp = gp
        self.axis_offset = 20
        self.left_axes = 1
        self.right_axes = 0
        self.GraphLayers = []
        #set up the left and bottom axes of the first layer
        select_graphpage(self.gp.Name,self.origin)
        select_layer(0,self.origin)
        show_axis(0,self.origin)#AXIS_BOTTOM
        show_axis(1,self.origin)#AXIS_LEFT

    def destroy(self):
        self.gp.Destroy()
    def new_GraphLayer(self,side='right'): #this needs fixing
        self.gp.Layers(0).Activate()
        
        assert(side=='right' or side=='left')
        if side == 'right':
            self.origin.Execute('layadd type:=rightY;')
            self.right_axes+=1
        elif side == 'left':
            self.origin.Execute('layadd type:=leftY;')
            self.left_axes+=1
        lay_index = self.gp.Layers.Count-1
        gl = self.gp.Layers(lay_index)
        gl.Name = 'Layer{}'.format(lay_index+1)
        select_graphpage(self.gp.Name,self.origin)
        select_layer(lay_index,self.origin)
        if side == 'right':
            offset = self.axis_offset*(self.right_axes-1)
            yaxis_pos_offset_right(offset,self.origin)
        elif side == 'left':
            offset = self.axis_offset*(self.left_axes-1)
            yaxis_pos_offset_left(offset,self.origin)
        self.fit_page_to_layers(self.gp)
        gl_wrapper = GraphLayer(self.gp,gl,self.origin)
        self.GraphLayers.append(gl_wrapper)
        return gl_wrapper
    def get_layers(self):
        gl_wrappers = [GraphLayer(self.gp,self.gp.Layers(i),self.origin) for i in range(self.gp.Layers.Count)]
        self.GraphLayers = gl_wrappers
        return gl_wrappers
    layers = property(get_layers)

class GraphLayer(Py2OriginC): #nned to make this a child
    def __init__(self,gp,gl,origin=None):
        Py2OriginC.__init__(self,origin)
        self.gp = gp
        self.gl = gl
        self.DataPlots = []
        self.y_axis_color_automatic()
    def destroy(self):
        self.gl.Destroy()
    def new_DataPlot(self,ws,x_i,y_i,plot_type='line'):
        #make datarange object
        dr = self.origin.NewDataRange()
        dr.Add('X',ws.ws,0,x_i,-1,x_i)
        dr.Add('Y',ws.ws,0,y_i,-1,y_i)
        #make plot
        if plot_type == 'line':
            dp = self.gl.AddPlot(dr,200)
        elif plot_type == 'scatter':
            self.gl.AddPlot(dr,201)
        elif plot_type == 'linesymb':
            self.gl.AddPlot(dr,202)
        dps = self.gl.DataPlots
        dp = dps[-1]
        dp_wrapper = DataPlot(self.gp,self.gl,dp,self.origin)
        self.DataPlots.append(dp_wrapper)
        return dp_wrapper
    def x_range(self,start,end):
        self.gl.Execute('layer.x.from = {};'.format(start))
        self.gl.Execute('layer.x.to = {};'.format(end))
        self.auto_increment()
        self.fit_page_to_layers(self.gp)
    def y_range(self,start,end):
        self.gl.Execute('layer.y.from = {};'.format(start))
        self.gl.Execute('layer.y.to = {};'.format(end))
        self.auto_increment()
        self.fit_page_to_layers(self.gp)
    def x_scale(self):
        pass
    def y_scale(self):
        pass
    def x_title(self,label):
        self.gl.Execute('label -xb {};'.format(label))
        self.fit_page_to_layers(self.gp)
    def y_title(self,label):
        self.gl.Execute('label -yl {};'.format(label))
        self.fit_page_to_layers(self.gp)
    def x_title_size(self,size):
        select_graphpage(self.gp.Name,self.origin) # calling originC functions
        select_layer(self.gl.Index,self.origin)
        xtitle_size(size,self.origin)
        self.fit_page_to_layers(self.gp)
    def y_title_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        ytitle_size(size,self.origin)
        self.fit_page_to_layers(self.gp)
    def x_label_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        xaxis_label_size(size,self.origin)
    def y_label_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        yaxis_label_size(size,self.origin)      
    def y_label_format(self,nformat):
        #1=scientific notation
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        yaxis_label_numeric_format(nformat,self.origin)        
    def x_major_tick_inc(self,increment):
        self.gl.Execute('layer.x.inc = {};'.format(increment))
        self.fit_page_to_layers(self.gp)
        self.fit_page_to_layers(self.gp)
    def y_major_tick_inc(self,increment):
        self.gl.Execute('layer.y.inc = {};'.format(increment))
        self.fit_page_to_layers(self.gp)
        self.fit_page_to_layers(self.gp)
    def y_axis_color_automatic(self):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        yaxis_color_automatic(self.origin)
        self.fit_page_to_layers(self.gp)
    def auto_increment(self,axis='X'):
        #using originc functions
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        if axis=='X':
            xsmart_axis_increment(self.origin)
        elif axis=='Y':
            ysmart_axis_increment(self.origin)
    def get_plots(self):
        dps = self.gl.DataPlots
        dp_wrappers = [DataPlot(self.gp,self.gl,d,self.origin) for d in dps]
        self.DataPlots = dp_wrappers
        return dp_wrappers
    plots = property(get_plots)
class DataPlot(Py2OriginC):
    def __init__(self,gp,gl,dp,origin=None):
        Py2OriginC.__init__(self,origin)
        self.gp = gp
        self.gl = gl
        self.dp = dp
    def destroy(self):
        self.dp.Destroy()
    def line_color(self,r,g,b):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_line_color(r,g,b,self.origin)
    def line_width(self,width):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_line_width(width,self.origin)
    def symb_type(self,s_type):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_marker_style(s_type,self.origin)
    def edge_color(self,r,g,b):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_marker_edge_color(r,g,b,self.origin)
    def face_color(self,r,g,b):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_marker_face_color(r,g,b,self.origin)
    def edge_width(self,width):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_marker_edge_width(width,self.origin)
# =============================================================================
# 
#     def range_obj(self):
#         cmd = 'range rr = !{};'.format(self.dp.Index+1)
#         self.gl.Execute(cmd)
#     def line_color(self,r,g,b):
#         self.range_obj()
#         cmd = 'set rr -cl color({},{},{});'.format(r,g,b)
#         self.gl.Execute(cmd)
#         self.gl.Execute(cmd)
#     def line_width(self,width):
#         self.range_obj()
#         cmd = 'set rr -w {};'.format(width)
#         self.gl.Execute(cmd)
#         self.gl.Execute(cmd)
#     def symb_type(self,s_type):
#         self.range_obj()
#         cmd = 'set rr -k {};'.format(s_type)
#         self.gl.Execute(cmd)
#     def edge_color(self,r,g,b):
#         self.range_obj()
#         cmd = 'ser rr -c color({},{},{});'.format(r,g,b)
#         self.gl.Execute(cmd)
#     def face_color(self,r,g,b):
#         self.range_obj()
#         cmd = 'set rr -cf color({},{},{});'.format(r,g,b)
#         self.gl.Execute(cmd)
#     def edge_width(self,width):
#         self.range_obj()
#         cmd = 'set rr -kh {};'.format(width)
#         self.gl.Execute(cmd)
# =============================================================================

        
'''
class GraphPage(Py2OriginC):
    def __init__(self,name,origin):
        Py2OriginC.__init__(self,origin)
        self.name = name
        self.GraphLayers = []
        #is there a way to capture the output from labtalk?
        make_graph(name,self.origin)
    def add_layer(self,position='right'):
        gl_index = len(self.graphLayers)
        gl = GraphLayer(gl_index,position)
        self.GraphLayers.append(gl)
        return gl
    def get_layer(self,index):
        return self.graphLayers[index]
    '''
# =============================================================================
# class GraphLayer(GraphPage):
#     def __init__(self,gl_index, position='right',origin=None):
#         Py2OriginC.__init__(self,origin)
#         self.DataPlots = []
#         self.position = position
#         self.gl_index = gl_index
#         if position == 'left':
#             add_xlinked_layer_left(self.origin)
#         elif position == 'right':
#             add_xlinked_layer_right(self.origin)
#     def add_dataPlot(self, WorkSheet_in, nCx, nCy, plot_type='scatter'):
#         select_layer(self.gl_index,self.origin)
#         dp_index = len(self.dataPlots)
#         dp = DataPlot(dp_index,WorkSheet_in, nCx, nCy, plot_type)
#         self.DataPlots.append(dp)
#         return dp
# class DataPlot(GraphLayer):
#     def __init__(self, dp_index, workSheet, nCx, nCy, plot_type='scatter',origin=None):
#         Py2OriginC.__init__(self,origin)
#         if plot_type == 'scatter':
#             make_scatter_plot(workSheet.name, nCx, nCy, self.origin)
#         elif plot_type == 'linesymb':
#             make_linesymb_plot(workSheet.name, nCx, nCy, self.origin)
#         elif plot_type == 'line':
#             make_line_plot(workSheet.name, nCx, nCy, self.origin)            
#         else:
#             pass
#         #probably need an index and activate function
#         self.dp_index = dp_index
#         self.xrange = None
#         self.yrange = None
#         self.xtitle = None
#         self.ytitle = None
#         self.xaxis_label_size = None
#         self.yaxis_label_size = None
#         self.xaxis_label_numeric_format = None
#         self.yaxis_label_numeric_format = None
#         self.xtitle_size = None
#         self.ytitle_size = None
#         self.xaxis_color = None
#         self.yaxis_color = None
#         self.yaxis_color_automatic = None
#         self.yaxis_pos_offset_left = None
#         self.yaxis_pos_offset_right = None
#         self.plot_marker_style = None
#         self.plot_marker_size = None
#         self.plot_marker_edge_color = None
#         self.plot_marker_edge_width = None
#         self.plot_marker_face_color = None
#         self.plot_line_style = None
#         self.plot_line_width = None
#     def activate(self):
#         select_layer(self.gl_index,self.origin)
#         select_plot(self.dp_index,self.origin)
#     def xrange(self,dFrom,dTo):
#         self.activate()
#         xrange(dFrom,dTo,self.origin)
#         self.xrange = (dFrom,dTo)
#     def yrange(self,dFrom,dTo):
#         self.activate()
#         yrange(dFrom,dTo,self.origin)
#         self.yrange = (dFrom,dTo)
#     def xtitle(self,strText):
#         self.activate()
#         xtitle(strText,self.origin)
#         self.xtitle = strText
#     def ytitle(self,strText):
#         self.activate()
#         ytitle(strText,self.origin)
#         self.ytitle = strText
#     def xaxis_label_size(self,dSize):
#         self.activate()
#         xaxis_label_size(dSize,self.origin)
#         self.xaxis_label_size = dSize
#     def yaxis_label_size(self,dSize):
#         self.activate()
#         yaxis_label_size(dSize,self.origin)
#         self.yaxis_label_size = dSize
#     def xaxis_label_numeric_format(self,nFormat):
#         self.activate()
#         xaxis_label_numeric_format(nFormat,self.origin)
#         self.xaxis_label_numeric_format = nFormat
#     def yaxis_label_numeric_format(self,nFormat):
#         self.activate()
#         yaxis_label_numeric_format(nFormat,self.origin)
#         self.yaxis_label_size = nFormat
#     def xtitle_size(self,dSize):
#         self.activate()
#         xtitle_size(dSize,self.origin)
#         self.xtitle_size = dSize
#     def ytitle_size(self,dSize):
#         self.activate()
#         ytitle_size(dSize,self.origin)
#         self.ytitle_size = dSize
#     def xaxis_color(self,nR,nG,nB):
#         self.activate()
#         xaxis_color(nR,nG,nB)
#         self.xaxis_color = (nR,nG,nB)
#     def yaxis_color(self,nR,nG,nB):
#         self.activate()
#         yaxis_color(nR,nG,nB,self.origin)
#         self.yaxis_color = (nR,nG,nB)
#     def yaxis_color_automatic(self):
#         self.activate()
#         yaxis_color_automatic(self.origin)
#         self.yaxis_color_automatic = True
#     def yaxis_pos_offset_left(self,dPosOffset):
#         self.activate()
#         yaxis_pos_offset_left(dPosOffset,self.origin)
#         self.yaxis_pos_offset_left = dPosOffset
#     def yaxis_pos_offset_right(self,dPosOffset):
#         self.activate()
#         yaxis_pos_offset_right(dPosOffset,self.origin)
#         self.yaxis_pos_offset_right = dPosOffset
#     def plot_marker_style(self,nMarkerStyle):
#         self.activate()
#         plot_marker_style(nMarkerStyle,self.origin)
#         self.plot_marker_style = nMarkerStyle
#     def plot_marker_size(self,dMarkerSize):
#         self.activate()
#         plot_marker_size(dMarkerSize,self.origin)
#         self.plot_marker_size = dMarkerSize
#     def plot_marker_edge_color(self,nR,nG,nB):
#         self.activate()
#         plot_marker_edge_color(nR, nG, nB,self.origin)
#         self.plot_marker_edge_color = (nR,nG,nB)
#     def plot_marker_edge_width(self,dWidth):
#         self.activate()
#         plot_marker_edge_width(dWidth,self.origin)
#         self.plot_marker_edge_width = dWidth
#     def plot_marker_face_color(self,nR,nG,nB):
#         self.activate()
#         plot_marker_face_color(nR,nG,nB,self.origin)
#         self.plot_marker_face_color = (nR,nG,nB)
#     def plot_line_style(self,nLineStyle):
#         self.activate()
#         plot_line_style(nLineStyle,self.origin)
#         self.plot_line_style = nLineStyle
#     def plot_line_width(self,dLineWidth):
#         self.activate()
#         plot_line_width(dLineWidth,self.origin)
#         self.plot_line_width = dLineWidth
# =============================================================================


    
   

#Acknowledgement:
#    This module contains pieces of code from python_to_originlab (author: jsbangsund, url: python_to_originlab https://github.com/jsbangsund/python_to_originlab)