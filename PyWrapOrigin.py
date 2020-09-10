
import OriginExt
import time
import os
import numpy as np
import pandas as pd
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
    #get path of this file
    folder_path = os.getcwd()
    #build the path to the origin C code
    code_path = os.path.join(folder_path,'Plotter.cpp')
    #run a labtalk command to load and compile the origin C code. See https://www.originlab.com/doc/LabTalk/ref/Run-obj
    origin.Execute('run.loadoc({},16);'.format(code_path))
    #wait for it to compile
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



# =============================================================================
# below are the python wrappers for functions written in originC
# They are called using Labtalk 
# See Plotter.cpp for reference
# =============================================================================

def load_plotter():
    pass
def make_graph(sGraphPageName,origin):
    origin.Execute('make_graph({});'.format(sGraphPageName))
def show_axis(nAxisType,origin):
    assert(nAxisType==0 or nAxisType==1 or nAxisType==2 or nAxisType==3)
    #nAxisType: 0=AXIS_BOTTOM, 1=AXIS_LEFT, 2=AXIS_TOP, 3=AXIES_RIGHT
    origin.Execute('show_axis({});'.format(nAxisType))
def select_graphpage(sGraphPageName,origin):
    origin.Execute('select_graphpage({});'.format(sGraphPageName))
def select_layer(nLayerInd,origin):
    origin.Execute('select_layer({});'.format(nLayerInd))
def select_plot(nPlotInd,origin):
    origin.Execute('select_plot({});'.format(nPlotInd))
def add_reflines_ver(lsReflines,origin):
    #input is a list of refline positions
    strReflines = ''
    for i,e in enumerate(lsReflines):
        if i > 0:
            strReflines += ' '
        strReflines += str(e)
    if len(strReflines) != 0:
        origin.Execute('add_reflines_ver({});'.format(strReflines))
def add_reflines_hor(strReflines,origin):
    origin.Execute('add_reflines_hor({});'.format(strReflines))
def refline_fill_ver(nRefLineIndex,nFillToIndex,nR,nG,nB,origin):
    origin.Execute('refline_fill_ver({},{},{},{},{});'.format(nRefLineIndex,nFillToIndex,nR,nG,nB))    
def refline_fill_hor(nRefLineIndex,nFillToIndex,nR,nG,nB,origin):
    origin.Execute('refline_fill_hor({},{},{},{},{});'.format(nRefLineIndex,nFillToIndex,nR,nG,nB))
def add_xlinked_layer_right(origin):
    origin.Execute('add_xlinked_layer_right();')
def add_xlinked_layer_left(origin):
    origin.Execute('add_xlinked_layer_left();')
def xrange(dFrom,dTo,origin):
    origin.Execute('xrange({},{});'.format(dFrom,dTo))
def yrange(dFrom,dTo,origin):
    origin.Execute('yrange({},{});'.format(dFrom,dTo))
def xsmart_axis_increment(origin):
    origin.Execute('xsmart_axis_increment();')
def ysmart_axis_increment(origin):
    origin.Execute('ysmart_axis_increment();')
def axis_rescale_type(nAxisType,nType,origin):
    origin.Execute('axis_rescale_type({},{});'.format(nAxisType,nType))
def axis_rescale(nAxisType,nRescale,origin):
    origin.Execute('axis_rescale({},{});'.format(nAxisType,nRescale))
def axis_rescale_margin(nAxisType,dResMargin,origin):
    origin.Execute('axis_rescale_margin({},{});'.format(nAxisType,dResMargin))
def xaxis_label_size(dSize,origin):
    origin.Execute('xaxis_label_size({});'.format(dSize))
def yaxis_label_size(dSize,origin):
    origin.Execute('yaxis_label_size({});'.format(dSize))
def xaxis_label_numeric_format(nFormat,origin):
    origin.Execute('xaxis_label_numeric_format({});'.format(nFormat))
def yaxis_label_numeric_format(nFormat,origin):
    origin.Execute('yaxis_label_numeric_format({});'.format(nFormat))
def xtitle(strText,origin):
    origin.Execute('xtitle({});'.format(strText))
def ytitle(strText,origin):
    origin.Execute('ytitle({});'.format(strText))
def xtitle_size(dSize,origin):
    origin.Execute('xtitle_size({});'.format(dSize))
def ytitle_size(dSize,origin):
    origin.Execute('ytitle_size({});'.format(dSize))
def xaxis_color(nR,nG,nB,origin):
    origin.Execute('xaxis_color({},{},{});'.format(nR,nG,nB))
def yaxis_color(nR,nG,nB,origin):
    origin.Execute('yaxis_color({},{},{});'.format(nR,nG,nB))
def yaxis_color_automatic(origin):
    origin.Execute('yaxis_color_automatic();')
def yaxis_pos_offset_left(dPosOffset,origin):
    origin.Execute('yaxis_pos_offset_left({});'.format(dPosOffset))
def yaxis_pos_offset_right(dPosOffset,origin):
    origin.Execute('yaxis_pos_offset_right({});'.format(dPosOffset))
def make_linesymb_plot(sWksName,nCx,nCy,origin):
    origin.Execute('make_linesymb_plot({},{},{});'.format(sWksName,nCx,nCy))
def make_line_plot(sWksName,nCx,nCy,origin):
    origin.Execute('make_line_plot({},{},{});'.format(sWksName,nCx,nCy))
def make_scatter_plot(sWksName,nCx,nCy,origin):
    origin.Execute('make_scatter_plot({},{},{});'.format(sWksName,nCx,nCy))
def plot_marker_style(nMarkerStyle,origin):
    origin.Execute('plot_marker_style({});'.format(nMarkerStyle))
def plot_marker_size(dMarkerSize,origin):
    origin.Execute('plot_marker_size({});'.format(dMarkerSize))
def plot_marker_edge_color(nR,nG,nB,origin):
    origin.Execute('plot_marker_edge_color({},{},{});'.format(nR,nG,nB))
def plot_marker_edge_width(dWidth,origin):
    origin.Execute('plot_marker_edge_width({});'.format(dWidth))
def plot_marker_face_color(nR,nG,nB,origin):
    origin.Execute('plot_marker_face_color({},{},{});'.format(nR,nG,nB))
def plot_line_color(nR,nG,nB,origin):
    origin.Execute('plot_line_color({},{},{});'.format(nR,nG,nB))
def plot_line_style(nR,nG,nB,origin):
    origin.Execute('plot_line_style({},{},{});'.format(nR,nG,nB))
def plot_line_width(dLineWidth,origin):
    origin.Execute('plot_line_width({});'.format(dLineWidth))

# =============================================================================
# PyWrapOrigin is a higher level class to operate on origin objects
# =============================================================================
class PyWrapOrigin():
    #This is the base class that the rest of the objects will inherit from
    def __init__(self):
        self.origin = None
    def connect(self):
        # Connect to Origin client
        origin = OriginExt.ApplicationSI()
        origin.Visible = origin.MAINWND_SHOW # Make session visible
        # Wait for origin to compile
        origin.Execute("sec -poc 3.5")
        time.sleep(3.5)
        #get path of this file
        folder_path = os.getcwd()
        #build the path to the origin C code
        code_path = os.path.join(folder_path,'Plotter.cpp')
        #run a labtalk command to load and compile the origin C code. See https://www.originlab.com/doc/LabTalk/ref/Run-obj
        origin.Execute('run.loadoc({},16);'.format(code_path))
        #wait for it to compile
        time.sleep(3.5)
        self.origin = origin
    def disconnect(self):
        self.origin.Exit()
    def save_project(self,project_name,full_path):
        # File ending is automatically added by origin
        project_name = project_name.replace('.opju','').replace('.opj','')
        self.origin.Execute("save " + os.path.join(full_path,project_name))
    def new_GraphPage(self,name):
        self.origin.CreatePage(3,name)
        gp = self.origin.GraphPages(name)
        gp = GraphPage(gp,self.origin)
        return gp
    def new_WorkSheet(self,ws_name,wb_name):
        ws = None
        if self.origin.WorksheetPages(wb_name) == None: #if workbook doesn't exist, create a new one
            self.origin.CreatePage(2,wb_name,'Origin')
        wb = self.origin.WorksheetPages(wb_name)
        for i in range(wb.Layers.Count): #find the ws matching ws_name
            x = wb.Layers(i)
            if x.Name == ws_name:
                ws = x
        if ws == None: #if worksheet doesn't exist, create a new one
            wb.Layers.Add()
            ws = wb.Layers(wb.Layers.Count-1)
            ws.Name = ws_name
        ws_wrapper = WorkSheet(wb,ws,self.origin)
        return ws_wrapper
class WorkSheet():
    #this class will make and destroy worksheets
    #it will also enable data transfer as a dataframe
    def __init__(self,wb,ws,origin):
        self.origin = origin
        self.wb = wb
        self.ws = ws
    def destroy(self):
        self.ws.Destroy()
    def get_name(self):
        return self.ws.Name
    def from_df(self, df, units = None):
        #input will be a pandas dataframe
        if units != None:
            assert(len(units) != len(df.columns))
        self.ws.Cols = len(df.columns)
        for i, c in enumerate(df.columns):
            col = self.ws.Columns(i)
            #col = self.ws.Columns(self.ws.Columns.Count-1)
            col.LongName = c
            col.SetData(np.float64(df[c])) #This needs some testing
            if units != None:
                col.Units = units[i]
            #col.Units = 'something' #need to implement adding units to the worksheet    
    name = property(get_name)
class GraphObjectBase():
    def __init__(self):
        pass
    def fit_page_to_layers(self,gp,origin):
        gp.Layers(0).Activate() #for some reason, this has to be done before calling the x-function
        origin.Execute('pfit2l margin:=tight;') #x-function to fit page to layers
class GraphPage(GraphObjectBase):
    def __init__(self,gp,origin):# pass in the actual object 
        GraphObjectBase.__init__(self)
        self.origin = origin
        self.gp = gp
        self.axis_offset = 20 #default axis_offset
        self.left_axes = 1
        self.right_axes = 0
        self.GraphLayers = []
        #show the left and bottom axes of the first layer because they not on by default
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
        self.fit_page_to_layers(self.gp,self.origin)
        gl_wrapper = GraphLayer(self.gp,gl,self.origin)
        self.GraphLayers.append(gl_wrapper)
        gl_wrapper.y_axis_color_automatic()#make color automatic
        return gl_wrapper
    def set_axis_offset(self,layer_index,offset):
        #needs originC code to offset position automatically
        #needs an OriginC function to auto offset axis position
        #OriginC function to auto update axis positions when one is changed
        pass
    def get_layers(self):
        gl_wrappers = [GraphLayer(self.gp,self.gp.Layers(i),self.origin) for i in range(self.gp.Layers.Count)]
        self.GraphLayers = gl_wrappers
        return gl_wrappers
    def get_name(self):
        return self.gp.Name
    
    layers = property(get_layers)
    name = property(get_name)
class GraphLayer(GraphObjectBase): 
    def __init__(self,gp,gl,origin):
        #new layer will be made whenever get_layer in GraphPage is called.
        #will get attributes stored in the internal origin objects on demand.
        GraphObjectBase.__init__(self)
        self.origin = origin
        self.gp = gp
        self.gl = gl
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
        return dp_wrapper
    def reflines_ver(self,refList):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        add_reflines_ver(refList,self.origin)
    def refline_fill_ver(self,nRefLineIndex,nFillToIndex,nR,nG,nB):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        refline_fill_ver(nRefLineIndex,nFillToIndex,nR,nG,nB,self.origin)
    def x_range(self,start,end):
        self.gl.Execute('layer.x.from = {};'.format(start))
        self.gl.Execute('layer.x.to = {};'.format(end))
        self.auto_increment()
        self.fit_page_to_layers(self.gp,self.origin)
    def y_range(self,start,end):
        self.gl.Execute('layer.y.from = {};'.format(start))
        self.gl.Execute('layer.y.to = {};'.format(end))
        self.auto_increment()
        self.fit_page_to_layers(self.gp,self.origin)
    def x_scale(self):
        pass
    def y_scale(self):
        pass
    def x_title(self,label):
        self.gl.Execute('label -xb {};'.format(label))
        self.fit_page_to_layers(self.gp,self.origin)
    def y_title(self,label):
        self.gl.Execute('label -yl {};'.format(label))
        self.fit_page_to_layers(self.gp,self.origin)
    def x_title_size(self,size):
        select_graphpage(self.gp.Name,self.origin) # calling originC functions
        select_layer(self.gl.Index,self.origin)
        xtitle_size(size,self.origin)
        self.fit_page_to_layers(self.gp,self.origin)
    def y_title_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        ytitle_size(size,self.origin)
        self.fit_page_to_layers(self.gp,self.origin)
    def x_label_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        xaxis_label_size(size,self.origin)
        self.fit_page_to_layers(self.gp,self.origin)
    def y_label_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        yaxis_label_size(size,self.origin)   
        self.fit_page_to_layers(self.gp,self.origin)
    def y_label_format(self,nformat):
        #1=scientific notation
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        yaxis_label_numeric_format(nformat,self.origin)  
        self.fit_page_to_layers(self.gp,self.origin)
    def x_major_tick_inc(self,increment):
        self.gl.Execute('layer.x.inc = {};'.format(increment))
        self.fit_page_to_layers(self.gp,self.origin)
    def y_major_tick_inc(self,increment):
        self.gl.Execute('layer.y.inc = {};'.format(increment))
        self.fit_page_to_layers(self.gp,self.origin)
    def y_axis_color_automatic(self):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        yaxis_color_automatic(self.origin)
        self.fit_page_to_layers(self.gp,self.origin)
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
class DataPlot(GraphObjectBase):
    def __init__(self,gp,gl,dp,origin):
        GraphObjectBase.__init__(self)
        self.origin = origin
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

#Acknowledgement:
#    This module contains pieces of code from python_to_originlab (author: jsbangsund, url: python_to_originlab https://github.com/jsbangsund/python_to_originlab)