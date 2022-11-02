# SAP Control Identify Tool
This tool is made based on this document (SAP GUI 7.60): https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/a020c8f8cfaf48ec9b579d5961889639.html?version=760.01

It is built using SAP GUI 7.50 and SAP ERP Server. 

I have never had access to try it using other resources. It may have bug and doesn't work. 

## How to use

#### Enabling Scripting on the Server Side 

Make sure your SAP Server enables Scripting

If it doesnt, you can follow this document: https://help.sap.com/docs/IRPA/8ecea00c1f854fd0a433c4aef5da1ea2/001675913cc54719930aa8197478dcde.html


#### Sample

- Start your SAP 
- Click Load button. It will try to load all object into Object TreeView. Each tree view node includes some basic information. 
- Select tree view node will add its Id to clipboard automatically and visualize the sap control in SAP GUI inside a red rectangle
- After select node, you can click Load Object Details button to see full details at Object Details TreeView. A sample document will be loaded also.
- Interaction groups:
  - Set value for properties with read-write access
  - Execute selected method with input paramaters and display in Result.

#### Video

https://www.youtube.com/watch?v=wJiXb_WmCG0