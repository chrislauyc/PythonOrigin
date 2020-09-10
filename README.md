# PyWrapOrigin
A python wrapper to OriginLab that allows plotting graphs in Originlab using python. 

The module utilizes Origin C, LabTalk, and OriginExt to enable plotting Origin graphs from the ground up, without relying on making any template. 

 Installation:
 
 Basic Usage:
 ```python
 from PyWrapOrigin import *
 pc = PyWrapOrigin()
 pc.connect() #this will open and compile originlab. Also it will compile the originC module.
 
 #to make a new worksheet
 ws = pc.new_WorkSheet('sheet0','book0')
 
 #transfer to sheet from a pandas dataframe
 ws.from_df(df)
 
 #to make a new graph
 gp = pc.new_GraphPage('Graph1')
 
 #to make new layers
 lay1 = gp.new_GraphLayer('right')
 lay2 = gp.new_GraphLayer('right')
 
 #to select a layer
 lay0 = gp.layers[0]
 
 #make a new data plot
 dp = lay0.new_DataPlot(ws,0,1,'scatter')
 
 #modify the plot
 lay0.y_title('Sine')
 lay0.y_title_size(20)
 dp.symb_type(2)
 dp.edge_color(0,0,255)
 dp.face_color(255,255,255)
 
 #set range
 lay0.y_range(-1,1)
 
 #add reflines
 lay0.reflines_ver([2,6,9])
 
 #fill reflines
 lay0.refline_fill_ver(0,1,255,255,0)
 
 #delete objects
 gp.destroy()
 ```
 Dependencies:
 
 OriginExt
 
 Useful Links:
 
