# VBA-CMMS
## VBA based CMMS application for maintenance management.

Visual basic for application (VBA) still considered one of the pratical tools for managing, orgnasing and automating tasks. alotought it might not be much know for engineering and heavy industry In this example I would highlight one of the possible applications of VBA in the industrial context. 

## Context.

This application was developped as a part of my messions at Nexans data company as a Process & Manufactruing engineer to schedule, planing as well as managing maintenance's orders and interventions' autorisations and material resources tracing for the hardware equipment used on the production lines. 

## Functional requirements. 

The software application must provide the ability to : 

- Requiest for intervention. 
- Intervention tracability. 
- Notifying the personel remotely. 
- Interventions archiving. 
- Maintenance KPI visualization. 
- Password protection to access the archived data. 
- Switching back and forth betwween different tabs. 
- Easy to use Graphical User Interface (GUI).  


## Data storage

There are many possibilities on how we can store the maintenance tasks data (Archiving) which depends mainly on the cmms use case, usually for decentralised cmms it will be required to access the data from anywhere in this case an SQL DB whill be more appropeite to store and retrieve the data. For this use case  since the solution is for an industrial antity I assummed that the cmms will be used locally and without excessive costs, maining it will not be connected to any external servers other than for notification puposes, and the the data will be stored in a typical Excel worksheets in a dedicated PC workstation. 

