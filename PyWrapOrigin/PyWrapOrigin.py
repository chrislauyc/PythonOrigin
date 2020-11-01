
import OriginExt
import importlib
import time
import os
import numpy as np
import pandas as pd
import pkg_resources
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

# Colors


def get_originC_path():
    return pkg_resources.resource_filename('PyWrapOrigin','OriginC/Plotter.cpp')
def color(name): #return colors as rgb values
    if name =='black':
        return (0,0,0)
    elif name == 'white':
        return (255,255,255)
    elif name == 'red':
        return (255,0,0)
    elif name == 'blue':
        return (0,0,255)
    elif name == 'green':
        return (0,255,0)
    elif name == 'cyan':
        return (0,255,255)
    elif name == 'magenta':
        return (255,0,255)
    elif name == 'yellow':
        return (255,255,0)
    elif name == 'dark yellow':
        return (128,128,0)
    elif name == 'navy':
        return (0,0,128)
    elif name == 'purple':
        return (128,0,128)
    elif name == 'wine':
        return (128,0,0)
    elif name == 'olive':
        return (0,128,0)
    elif name == 'dark cyan':
        return (0,128,128)
    elif name == 'royal':
        return (0,0,120)
    elif name == 'orange':
        return (255,128,0)
    elif name == 'violet':
        return (128,0,255)
    elif name == 'pink':
        return (255,0,128)
    elif name == 'lt gray':
        return (192,192,192)
    elif name == 'gray':
        return (128,128,128)
    elif name == 'lt yellow':
        return (255,255,128)
    elif name == 'lt cyan':
        return (128,255,255)
    elif name == 'lt magenta':
        return (255,128,255)
    elif name == 'dark gray':
        return (64,64,64)
    else:
        raise ValueError('Color name not found!')
def index_color(index): #return colors as rgb values
    colors = [
        (0,0,0),
        (255,0,0),
        (0,0,255),
        (0,255,0),
        (0,255,255),
        (255,0,255),
        (255,255,0),
        (128,128,0),
        (0,0,128),
        (128,0,128),
        (128,0,0),
        (0,128,128),
        (0,0,120),
        (255,128,0),
        (128,0,255),
        (255,0,128),
        (192,192,192),
        (128,128,128),
        (255,255,128),
        (128,255,255),
        (255,128,255),
        (64,64,64)
    ]
    return colors[index]
class PyWrapOrigin():
    #This is the base class that the rest of the objects will inherit from
    def __init__(self):
        self.origin = None
        self.graphpages = []
        self.worksheets = []
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
        code_path = os.path.join(folder_path,'OriginC','Plotter.cpp')
        #run a labtalk command to load and compile the origin C code. See https://www.originlab.com/doc/LabTalk/ref/Run-obj
        origin.Execute('run.loadoc({},16);'.format(get_originC_path()))
        # --------------
        # to troubleshoot whether the origin c code load properly, run the following commands in labtalk
        # err = run.loadoc("E:\OneDrive - University of Utah\Anderson's lab\OriginC\Py2OriginC\OriginC\Plotter.cpp",16);
        # type $(err);
        # --------------
        #wait for it to compile
        time.sleep(3.5)
        self.origin = origin
    def disconnect(self):
        self.origin.Exit()
        self.origin = None
    def save_project(self,project_name,full_path):
        # File ending is automatically added by origin
        project_name = project_name.replace('.opju','').replace('.opj','')
        self.origin.Execute("save " + os.path.join(full_path,project_name))
    def new_GraphPage(self,name):
        self.origin.CreatePage(3,name)
        gp = self.origin.GraphPages(name)
        gp = GraphPage(gp,self.origin)
        self.graphpages.append(gp)
        return gp
    def new_WorkSheet(self,ws_name,wb_name):
        ws = None
        if self.origin.WorksheetPages(wb_name) == None: #if workbook doesn't exist, create a new one
            self.origin.CreatePage(2,wb_name,'Origin')
        wb = self.origin.WorksheetPages(wb_name)
        for i in range(wb.Layers.Count): #find the ws matching ws_name
            x = wb.Layers(i)
            if x.Name.lower() == ws_name.lower():
                ws = x
        if ws == None: #if worksheet doesn't exist, create a new one
            wb.Layers.Add()
            ws = wb.Layers(wb.Layers.Count-1)
            ws.Name = ws_name
        ws_wrapper = WorkSheet(wb,ws,self.origin)

        self.worksheets.append(ws_wrapper) #storing the worksheet
        return ws_wrapper
    def get_WorkSheet(self,ws_name,wb_name):
        return self.new_WorkSheet(ws_name,wb_name)
class WorkSheet():
    #this class will make and destroy worksheets
    #it will also enable data transfer as a dataframe

    #need a way to obtain worksheet by name
    def __init__(self,wb,ws,origin):
        self.origin = origin
        self.wb = wb
        self.ws = ws
    def destroy(self):
        self.ws.Destroy()
    def get_name(self):
        return '[{}]{}'.format(self.wb.Name,self.ws.Name)
    def from_df(self, df, units = None):
        #input will be a pandas dataframe
        if units != None:
            assert(len(units) != len(df.columns))
        self.ws.Cols = len(df.columns)
        for i, c in enumerate(df.columns):
            col = self.ws.Columns(i)
            #col = self.ws.Columns(self.ws.Columns.Count-1)
            col.LongName = c
            # col.SetData(np.float64(df[c])) #This needs some testing
            col.SetData(np.array(df[c]))
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
        yaxis_color_automatic(self.origin)
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

        #need to autoscale
        GraphObjectBase.__init__(self)
        self.origin = origin
        self.gp = gp
        self.gl = gl
        self.DataPlots = []
    def destroy(self):
        self.gl.Destroy()
    # def new_DataPlot(self,ws,x_i,y_i,plot_type='line'):
    #     #make datarange object
    #     dr = self.origin.NewDataRange()
    #     dr.Add('X',ws.ws,0,x_i,-1,x_i)
    #     dr.Add('Y',ws.ws,0,y_i,-1,y_i)
    #     #make plot
    #     if plot_type == 'line':
    #         dp = self.gl.AddPlot(dr,200)
    #     elif plot_type == 'scatter':
    #         self.gl.AddPlot(dr,201)
    #     elif plot_type == 'linesymb':
    #         self.gl.AddPlot(dr,202)
    #     dps = self.gl.DataPlots
    #     dp = dps[-1]
    #     dp_wrapper = DataPlot(self.gp,self.gl,dp,self.origin)
    #     return dp_wrapper
    def new_DataPlot(self,ws,x_i,y_i,plot_type='line'):
        #make plot
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        if plot_type == 'line':
            make_line_plot(ws.name,x_i,y_i,self.origin)
        elif plot_type == 'scatter':
            make_scatter_plot(ws.name,x_i,y_i,self.origin)
        elif plot_type == 'linesymb':
            make_linesymb_plot(ws.name,x_i,y_i,self.origin)
        dps = self.gl.DataPlots
        dp = dps[-1]
        # print(type(dps))
        index = len(dps)-1
        # for dp in dps:
        #     print(dp.Index)
        # print('-------------index: {}-----------'.format(index))
        dp_wrapper = DataPlot(self.gp,self.gl,dp,index,plot_type,self.origin)
        dp_wrapper.default_settings()
        self.DataPlots.append(dp_wrapper)
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
class DataPlot(GraphObjectBase):
    def __init__(self,gp,gl,dp,index,plot_type,origin):
        GraphObjectBase.__init__(self)
        self.origin = origin
        self.gp = gp
        self.gl = gl
        self.dp = dp
        self.index = index
        self.plot_type = plot_type
    def destroy(self):
        self.dp.Destroy()
    def default_settings(self):
        if self.plot_type == 'line':
            self.line_width(1)
        elif self.plot_type == 'scatter':
            self.symb_type(2)
            self.symb_size(3)
            self.edge_color(*index_color(self.gl.Index))
            self.face_color(*color('white'))
        elif self.plot_type == 'linesymb':
            self.symb_type(2)
            self.symb_size(3)
            self.edge_color(*index_color(self.gl.Index))
            self.face_color(*color('white'))
    def line_color(self,r,g,b):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_line_color(r,g,b,self.origin)
    def line_width(self,width):
        # print(self.gp.Name,self.gl.Index,self.index)
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.index,self.origin)
        plot_line_width(width,self.origin)
    def symb_type(self,s_type):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_marker_style(s_type,self.origin)
    def symb_size(self,size):
        select_graphpage(self.gp.Name,self.origin)
        select_layer(self.gl.Index,self.origin)
        select_plot(self.dp.Index,self.origin)
        plot_marker_size(size,self.origin)
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
