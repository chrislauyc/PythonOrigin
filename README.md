# Py2OriginC
An API to allow python to call Origin C functions for plotting in Origin


Automating plotting in origin are often done using labtalk APIs. It is often difficult and confusing as labtalk documentations are not complete and labtalk lacks full access to origin internal objects. These problems can be solved by using Origin C code, which is very well documented and is full of functionalities. This repo contains the Origin C/C++ code to allow Origin C functions to be called through labtalk commands. A python module will then communicate with Origin through the labtalk commands. 
