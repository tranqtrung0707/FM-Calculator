# FM-Calculator
An interactive spreadsheet built in Excel/VBA to automate SOA Exam FM calculations, including constant &amp; variable annuities, amortization schedules, and duration/convexity.

## Annuities
Calculate annuity variables for ordinary annuities (equal payment each period), geometric annuities (payment increases by a fixed percentage each period), and arithmetic annuities (payment increases by a fixed amount each period).
Instructions: Select the mode, enter values for all but one input, then click CPT to calculate the blank cell's value. Calculations will not be made if there are 0 or multiple blank cells.
The computation results in a net present value of 0, so the cash flows cannot all have the same sign.
* Mode: BGN indicates payment made at the beginning of each period, END indicates payment made at the end of each period.
* N: Number of payment periods. "inf" (without quotes) denotes perpetuity. In perpetuity mode, FV is set to 0 regardless of user input.
* I: Interest rate per payment period (in percentage, i.e. value of 10 is understood as 10%)
* PV: Present value
* PMT: Fixed payment per period
* FV: Future value
* G: Growth rate of the payments (in percentage form, e.g. input of 5 is understood as 5%)
* PMT1: First payment of a geometric annuity
* P: First payment of an arithmetic annuity
* Q: The amount each subsequent payment is increased by
