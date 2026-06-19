# FM-Calculator
An interactive spreadsheet built in Excel/VBA to automate SOA Exam FM calculations, including constant &amp; variable annuities, amortization schedules, and duration/convexity.

## Annuities
Calculate annuity variables for ordinary annuities (equal payment each period), geometric annuities (payment increases by a fixed percentage each period), and arithmetic annuities (payment increases by a fixed amount each period).

<img width="619" height="240" alt="image" src="https://github.com/user-attachments/assets/bffe7774-b554-4a67-a5c4-4e49d28c77be" />

Instructions: Select the mode (BGN or END), enter values for all but one input, then click CPT to calculate the blank cell's value. Calculations will not be made if there are 0 or multiple blank cells.

The computation results in a net present value of 0, so the cash flows cannot all have the same sign.

The inputs are:
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

## Amortization
<img width="843" height="437" alt="image" src="https://github.com/user-attachments/assets/c8fdfe31-39f9-4e59-bda0-fe22fb6a0ed0" />

Enter the number of payments, interest per period (in percentage), loan amount and level payment. Click *Create Amortization Table* to create the amortization schedule and calculate the remaining balance.

Click *Clear Amortization Table* to clear the amortization schedule.

## Duration & Convexity
Calculate duration and convexity for a series of cash flows given by the boxed table.

<img width="484" height="144" alt="image" src="https://github.com/user-attachments/assets/57fad81c-3828-4ef5-8ab2-eda21718b034" />

Each cash flow value has a corresponding frequency, which is the number of periods that cash flow value is received.

The first cash flow is received at time 0.

Interest rate is in percentage, e.g. input of 10 is understood as 10%.

The maximum number of cash flow values is 50, i.e. it must be inputed inside the box.
