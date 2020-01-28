# PERT_GANTT
Creates a Gantt diagram with PERT estimates.

All in one diagram, it shows this information for each task:
Opmistic time (o), Most-likely time (m), Pessimistic time (p), Current work done, projected remaining work done, relations to other tasks

A VBA script runs on an Excel table, and shows the information visually. 

The times estimated, o, m, and p, are consistent with the PERT system, which is a weighted everage of the three times. Here is the formula:
  Expected time = (o + 4m + p)

A combobox from the chart sheet allows you to view the chart with three-different time bases: optimistic, most-likely, and pessimistic

Calculated on a percentage basis.
