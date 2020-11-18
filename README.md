# PyWrapOrigin
OriginPro is a graphing and analysis software used by many scientists. Sometimes making many plots by hand can be tedious and is necessary for automation. The OriginPro software is packaged with its own programming languages such as Origin C and LabTalk. But they are difficult for users to learn and implement.

PyWrapOrigin is a python wrapper to Origin that allows plotting graphs in Origin using python.

The module utilizes Origin C, LabTalk, and OriginExt to enable plotting Origin graphs from the ground up, without relying on using any template.

Installation:
 -----
Clone the repo:

 ``
 git clone https://github.com/chrislauyc/PyWrapOrigin.git
 ``

 In the cloned directory:

 ``
 python setup.py sdist
 ``

 ``
 pip install .
 ``

 Usage:
 -----
Import Library
 ```python
 from PyWrapOrigin.PyWrapOrigin import PyWrapOrigin

 ```
 Connect to Originlab. This will take several seconds to open Origin and import the OriginC code.
 ```python
 pwo = PyWrapOrigin()
 pwo.connect()
 ```
 Some arbitrary data in a pandas dataframe to feed into Origin
 ```python
 import numpy as np
 import pandas as pd
 x1 = np.linspace(0,2)
 y1 = np.exp(x1)
 y2 = np.exp(2*x1)

 data = {
     'x1':x1,
     'y1':y1,
     'y2':y2
 }
 df = pd.DataFrame(data)
 ```
 Create a new worksheet and send df into Origin
 ```python
 ws = pwo.new_WorkSheet('sheet1','book0')
 ws.from_df(df)
 ```
 Create a new graph page
 ```python
 gp = pwo.new_GraphPage('Graph1')
 ```
 A new graph pages already has 1 layer. To add a new layer on the right:
 ```python
 gp.new_GraphLayer('right')
 ```
 The layers can be accessed through the layers attribute of the graph page object.
 ```python
 lay0 = gp.layers[0]
 lay1 = gp.layers[1]
 ```
 Create new data plots in each layer.
 ```python
 dp0 = lay0.new_DataPlot(ws,0,1,'scatter')
 dp1 = lay1.new_DataPlot(ws,0,2,'scatter')
 ```
 Layer and data plot settings.
 ```python
 # Settings of the first layer
 lay0.y_title('Exp(x)')
 lay0.y_title_size(20)
 lay0.x_title('x')
 lay0.x_title_size(20)
 # first data plot
 dp0.symb_type(2) #circle
 dp0.edge_color(0,0,255) #RGB
 dp0.face_color(255,255,255) #RGB
 # Second layer
 lay1.y_title('Exp(2x)')
 lay1.y_title_size(20)
 # Second data plot
 dp1.symb_type(2) #circle
 dp1.edge_color(255,0,0) #RGB
 dp1.face_color(255,255,255) #RGB
 ```
 Figure plotted.

![Origin Plot Example]()

Reference lines can also be added to the figure
 ```python
 # List of x axis values as the reference line positions
 lay0.reflines_ver([0,1,2])
 # Can fill the space between two reflines with colors
 lay0.refline_fill_ver(0,1,128,172,242)
 ```
![Origin Plot With Ref Lines]()

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
