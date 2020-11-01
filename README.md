# PyWrapOrigin
A python wrapper to OriginLab that allows plotting graphs in OriginLab using python.

The module utilizes Origin C, LabTalk, and OriginExt to enable plotting Origin graphs from the ground up, without relying on making any template.

 Manual Installation:
 -----
Clone the repo:

 ``
 git clone https://github.com/chrislauyc/PyWrapOrigin.git
 ``

 In the cloned directory:

 ``
 python setyp.py sdist
 ``

 ``
 pip install .
 ``

 Usage:
 -----
 ```python
 import PyWrapOrigin
 po = PyWrapOrigin.PyWrapOrigin()
 po.connect() #this will open and compile originlab. Also it will compile the originC module.

 #to make a new worksheet
 ws = po.new_WorkSheet('sheet0','book0')

 #transfer to sheet from a pandas dataframe
 ws.from_df(df)

 #to make a new graph
 gp = po.new_GraphPage('Graph1')

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
 -----

Python version tested: Python 3.7

Operating system tested: Windows 10

Origin version tested: OriginPro 2018

Requires ``OriginExt``, ``numpy``, and  ``pandas``

Requires OriginLab installed

Requires win32com python package (https://anaconda.org/anaconda/pywin32)

 Testing
 -----

Development
-----

Feel free to open issues and pull requests

 Useful Links:
 -----
https://github.com/jsbangsund/python_to_originlab
