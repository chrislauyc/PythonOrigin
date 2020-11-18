# PyWrapOrigin
Origin is a graphing and analysis software used by many scientists. Making many plots by hand can be tedious, so sometimes it's better to automate the process. The Origin software is packaged with its own programming languages such as Origin C and LabTalk. But they are difficult for users to learn and implement. Origin also provides python APIs such as the automation server, PyOrigin, Originpro and OriginExt, which intend to provide access to Origin from python. However, none of these tools support full plotting functionalities, and it is often necessary to first create a graph template, then send data from python to be plotted. This adds time and inconvience in plotting. 

PyWrapOrigin is a python wrapper to Origin C, LabTalk, and OriginExt that allows plotting graphs in Origin from python possible and easy to do.

This module enables plotting Origin graphs from the ground up, without relying on using any template.

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

![Origin Plot Example](/imgs/plot1.JPG)

Reference lines can also be added to the figure
 ```python
 # List of x axis values as the reference line positions
 lay0.reflines_ver([0,1,2])
 # Can fill the space between two reflines with colors
 lay0.refline_fill_ver(0,1,128,172,242)
 ```
![Origin Plot With Ref Lines](/imgs/plot1_ver_fill.JPG)

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
