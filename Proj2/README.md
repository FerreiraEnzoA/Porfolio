# FerreiraEnzoA
:computer: Problem solver and data-driven professional <br />
:email: Email: <a href="mailto:FerreiraEnzoA@gmail.com">FerreiraEnzoA@gmail.com</a> <br />
:briefcase: LinkedIn: [FerreiraEnzoA](https://www.linkedin.com/in/ferreiraenzoa/) <br /> <br /> <br />


<p align="center">
:point_right: <a href="/Portfolio">MAIN PAGE</a> || <a href="/Portfolio/Archive">ARCHIVE</a> || <a href="/Portfolio/About">ABOUT</a> :point_left:
</p>
<br /> <br /> <br />


# <img src="../images/Python.png" alt="R" width="50"/> Project 2: Automation of Reconfirmation <br />
This project was inspired by difficulties in reconfirming a company's sales to a specific customer, as the task took hours to complete. In this context, the objective is to reduce both the reconfirmation time and the possible human errors in the analysis of extensive spreadsheets as much as possible.

![](../images/proj2.png)

The Customer Service sends a confirmation `conf_cs.xlsx` to the customer, then the customer replies with his own confirmation `conf_customer.xlsx`, so an analysis of correspondence or non-correspondence between these documents is necessary. In this way, this automation project proposes to analyze both (partially, since the final sales decisions with the necessary corrections depend on the Customer Service), generate a file for a possible reconfirmation `reconfirmation.xlsx` and indicate the warnings (items/quantity/arrival date) as output. CPC refers to Customer Product Code.

Primarily, I used the pandas library to manipulate structured data and, secondarily, openpyxl library to read/write Excel 2010 .xlsx files.

You can both get the [source code](https://github.com/FerreiraEnzoA/Portfolio/tree/main/Proj2)

![](../images/proj2_diagram.png)

An issue to be solved is when there is more than the same item in both worksheets, as the output would be a false error if the quantity or arrival date differed, despite the fact that the worksheet is generated perfectly, so it is a wrong warning.
