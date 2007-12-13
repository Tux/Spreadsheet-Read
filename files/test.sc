# This model file is a sample of a SquirrelCalc Spreadsheet.
#
# (c) 1986 Minihouse Research

dimension 28 {of 28 rows} 10 {of 10 columns}

global frame 2
global display 2
global format [F2L]
width c1 16
width c2 2
width c3 9
width c4 4
width c5 23
width c6 2
text na "N/A"
r0c0 = "sample1.sc"
lock r0c0
lock r0c1
lock r0c2
r0c3 = "Maximum Loan Amount Table"
lock r0c3
lock r0c4
lock r0c5
r0c6 = date(today)
lock r0c6
lock r0c7
r1c0 = [G2*] "-"
lock r1c0
r1c1 = [G2*] "-"
lock r1c1
r1c2 = [G2*] "-"
lock r1c2
r1c3 = [G2*] "-"
lock r1c3
r1c4 = [G2*] "-"
lock r1c4
r1c5 = [G2*] "-"
lock r1c5
r1c6 = [G2*] "-"
lock r1c6
r1c7 = [G2*] "-"
lock r1c7
r2c0 = "          Assumptions"
lock r2c0
lock r2c1
lock r2c2
lock r2c3
lock r2c4
r2c5 = "          Results"
lock r2c5
lock r2c6
lock r2c7
r3c0 = [G2*] "-"
lock r3c0
r3c1 = [G2*] "-"
lock r3c1
r3c2 = [G2*] "-"
lock r3c2
r3c3 = [G2*] "-"
lock r3c3
r3c4 = [G2*] "-"
lock r3c4
r3c5 = [G2*] "-"
lock r3c5
r3c6 = [G2*] "-"
lock r3c6
r3c7 = [G2*] "-"
lock r3c7
r4c0 = "Monthly Income"
lock r4c0
lock r4c1
r4c2 = ":"
lock r4c2
r4c3 = [$2L] 3500
lock r4c4
r4c5 = "Maximum Loan Amount"
lock r4c5
r4c6 = ":"
lock r4c6
r4c7 = r25c1
lock r4c7
lock r5c0
lock r5c1
lock r5c2
lock r5c3
lock r5c4
lock r5c5
lock r5c6
lock r5c7
r6c0 = "% of Income towards Repay"
lock r6c0
lock r6c1
r6c2 = ":"
lock r6c2
r6c3 = [%1L] 0.3
lock r6c4
r6c5 = "Affordable House"
lock r6c5
r6c6 = ":"
lock r6c6
r6c7 = r4c7+r15c3
lock r6c7
lock r7c0
lock r7c1
lock r7c2
lock r7c3
lock r7c4
lock r7c5
lock r7c6
lock r7c7
r8c0 = "Percentage of Loan Payment"
lock r8c0
lock r8c1
lock r8c2
lock r8c3
lock r8c4
r8c5 = "Required Down Payment"
lock r8c5
r8c6 = ":"
lock r8c6
r8c7 = r6c7*r17c3
lock r8c7
r9c0 = "towards Tax, Ins, Assmnts"
lock r9c0
lock r9c1
r9c2 = ":"
lock r9c2
r9c3 = [%1L] 0.35
lock r9c4
lock r9c5
lock r9c6
lock r9c7
lock r10c0
lock r10c1
lock r10c2
lock r10c3
lock r10c4
r10c5 = "Maximum Monthly Paymnt"
lock r10c5
r10c6 = ":"
lock r10c6
r10c7 = r4c3*r6c3
lock r10c7
r11c0 = "Term of the Loan in Years"
lock r11c0
lock r11c1
r11c2 = ":"
lock r11c2
r11c3 = [I2L] 29
lock r11c4
lock r11c5
lock r11c6
lock r11c7
lock r12c0
lock r12c1
lock r12c2
lock r12c3
lock r12c4
r12c5 = "Max. Loan Paymnt/Month"
lock r12c5
r12c6 = ":"
lock r12c6
r12c7 = r10c7/(1+r9c3)
lock r12c7
r13c0 = "Interest of the Loan"
lock r13c0
lock r13c1
r13c2 = ":"
lock r13c2
r13c3 = [%2L] 0.1475
lock r13c4
lock r13c5
lock r13c6
lock r13c7
lock r14c0
lock r14c1
lock r14c2
lock r14c3
lock r14c4
lock r14c5
lock r14c6
lock r14c7
r15c0 = "Available for Down Payment"
lock r15c0
lock r15c1
r15c2 = ":"
lock r15c2
r15c3 = [$2L] 25000
lock r15c4
lock r15c5
lock r15c6
lock r15c7
lock r16c0
lock r16c1
lock r16c2
lock r16c3
lock r16c4
lock r16c5
lock r16c6
lock r16c7
r17c0 = "Required down Payment"
lock r17c0
lock r17c1
r17c2 = ":"
lock r17c2
r17c3 = [%1L] 0.1
lock r17c4
r17c5 = "Adapted from: \"VisiCalc Home and"
lock r17c5
lock r17c6
lock r17c7
lock r18c0
lock r18c1
lock r18c2
lock r18c3
lock r18c4
r18c5 = "               Office Companion\""
lock r18c5
lock r18c6
lock r18c7
r19c0 = "Payments per Year"
lock r19c0
lock r19c1
r19c2 = ":"
lock r19c2
r19c3 = [I2L] 12
lock r19c4
r19c5 = date(31320,"Date created: %d-%h-%Y")
lock r19c5
lock r19c6
lock r19c7
lock r20c0
lock r20c1
lock r20c2
lock r20c3
lock r20c4
lock r20c5
lock r20c6
lock r20c7
r21c0 = "  Workspace"
lock r21c0
lock r21c1
lock r21c2
lock r21c3
lock r21c4
lock r21c5
lock r21c6
lock r21c7
r22c0 = "Calc 1:"
lock r22c0
r22c1 = r4c3*r6c3/(1+r9c3)*r19c3
lock r22c1
lock r22c2
lock r22c3
lock r22c4
lock r22c5
lock r22c6
lock r22c7
r23c0 = "Calc 2:"
lock r23c0
r23c1 = 1/(r13c3/r19c3+1)^(r11c3*r19c3)
lock r23c1
lock r23c2
lock r23c3
lock r23c4
lock r23c5
lock r23c6
lock r23c7
r24c0 = "Calc 3:"
lock r24c0
r24c1 = 1-r23c1
lock r24c1
lock r24c2
lock r24c3
lock r24c4
lock r24c5
lock r24c6
lock r24c7
r25c0 = "Calc 4:"
lock r25c0
r25c1 = r22c1/r13c3*r24c1
lock r25c1
lock r25c2
lock r25c3
lock r25c4
lock r25c5
lock r25c6
lock r25c7
r26c0 = ""
r26c1 = " "
r27c0 = " "
r27c1 = ""
goto r0c0
model on
