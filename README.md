# ETABS DRIFT CHECKING PROGRAMS

### Table of Contents
* [Purpose](#purpose)
* [Technologies](#technologies)
* [Project Status](#project-status)
* [Project Background](#project-background)
* [Drift Table Demo](#table-demo)
* [API Tool Demos](#API-demo)


### Purpose
These programs aim to automate tediuos formatting of drift output from ETABS for structural design of drift-controlled lateral systems.

### Technologies
 - Pandas
 - ETABS API
 - PYQT5

### Project Status
 - Drift calculations are stable and relatively bug-free. 
 - Torsion calculation seems to be buggy, not sure how to more appropriately calculate torsion ratio
 - UI and functionality could definitely use improvements and additions

### Project Background
This collection of programs was initially developed on a hospital project to speed up the size optimization for lateral system elements. The drift outputs from ETABS can be difficult to use and slow to format so these programs attempt to streamline the process and automate the tedious formatting of ETABS output.

### Drift Table Demo
![Drift Table Tool Demo](demos/driftTable_demo.gif)


### API Tool Demos
![Open ETABS Model and Read Drifts Demo](demos/APItool_open_read_demo.gif) <br />
![Formatted Excel Results Demo](demos/APItool_excelResults_demo.gif) <br />
![Use APItool Interactively with ETABS Demo](demos/APItool_interactiveUse_demo.gif)
