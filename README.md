# DrawioGen
Tools for generating XML Diagram in Draw.io. Diagram generated is specifically for Metro/IPRAN/Router network topology.

A straightforward tool to generate draw io network topology diagram.


<img width="433" height="333" alt="image" src="https://github.com/user-attachments/assets/cf8c3c87-12bc-4071-b76d-66443c0fd628" />

This is built for ease of generating a topology without the hassle to draw the topology manually.

##Data required (in single excel file):
**1. Network Element (NE) data (Hostname, IP loopback, NE node type, Site ID)**
  <img width="683" height="133" alt="image" src="https://github.com/user-attachments/assets/5534e869-fae9-4bb7-832b-508115c3e7ca" />

 node_type includes: Metro Router (ME), BNG, CSR, OLT, BTS towers, others
 
**2. Link data (connection between NE)**

<img width="512" height="60" alt="image" src="https://github.com/user-attachments/assets/886d8f5f-0bd5-45b3-96ef-d7ab93f8ac28" />

This specifies node a to node b connection. Which will be connected on the topology using a straight line.

##Instruction:
1. Put excel file for input in the TopoExcel directory. Template file available in the directory.
2. Output will be generated inside OutputTopo directory
3. Run the python source code and follow instructions.

##**Release Notes:**

_release version v1.00:_
  1. Single page diagram output.
  2. Input using .xlsx file.
  3. UI Console (Error checking and data re-entry after finished 1 task)
   
