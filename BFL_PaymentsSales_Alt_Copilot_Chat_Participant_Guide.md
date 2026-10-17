# Copilot Chat for Excel — Participant Exercise Guide
### Formulas & Pivot Tables for Payments Field Sales
**Bajaj Finance Limited | M365 Copilot Training**
**File:** `BFL_PaymentsSales_Copilot_Demo.xlsx`
**Mode:** Copilot Chat (no Microsoft 365 Copilot licence required). Works with Chat on a laptop, or on your phone alongside Excel on a laptop.

---

## Before You Begin

### What You Need
- **Excel** open with `BFL_PaymentsSales_Copilot_Demo.xlsx`
- **Copilot Chat**, any of:
  - `https://m365.cloud.microsoft/chat` (work account)
  - The **Copilot** button in Microsoft Edge
  - The **M365 Copilot app** or Teams on your phone
- If you are on one laptop, put the two windows side by side (**Win + ←** and **Win + →**).

### How This Session Works
Copilot Chat is **not connected to your workbook.** It writes the formula or the steps, and **you** paste them into Excel.

| Step | Where | What You Do |
|------|-------|-------------|
| **1. Ask** | Copilot Chat | Paste the prompt; get the formula or the steps |
| **2. Apply** | Excel | Paste into the cell, press Enter, fill down |

**The sheet:** headers are in **row 2**, and data runs from **row 3 to row 27** (25 merchants, as of 30 Sep 2026). Columns **N–Q are empty**. *(Column M is a deliberate blank spacer; it matters in Exercise 6.)*

---

## Step 0 — Set the Context (do this once)

Paste this into Copilot Chat **first**:

```
I'm working in Excel on a merchant onboarding tracker. You can't see the file,
so here is the layout:

Headers are in row 2. Data runs from row 3 to row 27 (25 rows).
The data is as of 30 September 2026.

A  Merchant ID          (text, e.g. MER-4101)
B  Merchant Name        (text)
C  Category             (text, e.g. Pharmacy, Restaurant)
D  Area                 (Pune West, Pune East, Mumbai Andheri, Mumbai Dadar, Nashik)
E  Product              (UPI QR, UPI QR + Soundbox, Android POS, Payment Link)
F  FOS                  (field officer name)
G  Lead Date            (date)
H  Installed On         (date; blank if not installed yet)
I  Install SLA          (number of days allowed from lead date)
J  First Txn On         (date; blank if no transaction yet)
K  Sep TPV (₹)          (number)
L  Sep Txn Count        (number)
M  (blank spacer — ignore it)
N  Days to Install      (empty — I'll add this)
O  Install SLA Status   (empty — I'll add this)
P  Activation Status    (empty — I'll add this)
Q  Usage Band           (empty — I'll add this)

For the rest of this conversation: when I ask for a formula, give me the
formula for row 3, ready to paste and fill down. Formula on its own line,
then a one-line explanation. No long preamble.
```

> 💡 **Keep this block.** Save it in a note on your phone and adapt it for any tracker you use.

---

## Exercise 1 — Days to Install (Column N)

**Ask:**
```
Formula for N3: days from G3 to H3. If H3 is blank, return blank.
```
**Expected:** `=IF(H3="","",H3-G3)`

**Apply:** paste into **N3**, press Enter, then double-click the fill handle to fill down to **N27**. If you see dates instead of numbers, set the column format to **General**.

✅ **Check:** row 3 shows **3**. Three rows are blank.

---

## Exercise 2 — Install SLA Status (Column O)

**Ask:**
```
Formula for O3: "Not Installed" if H3 is blank; "Breach" if N3 is greater
than I3; otherwise "Within SLA".
```
**Expected:** `=IF(H3="","Not Installed",IF(N3>I3,"Breach","Within SLA"))`

✅ **Check:** 10 Breach, 12 Within SLA, 3 Not Installed.

---

## Exercise 3 — Activation Status (Column P)

**Ask:**
```
Formula for P3: "Not Installed" if H3 is blank; "Inactive" if J3 is blank;
otherwise "Active".
```
**Expected:** `=IF(H3="","Not Installed",IF(J3="","Inactive","Active"))`

✅ **Check:** 17 Active, 5 Inactive, 3 Not Installed.

---

## Exercise 4 — Usage Band (Column Q)

**Ask:**
```
Formula for Q3: blank if P3 is not "Active"; otherwise "High" if K3 is
200000 or more, "Medium" if 50000 or more, else "Low".
```
**Expected:** `=IF(P3<>"Active","",IF(K3>=200000,"High",IF(K3>=50000,"Medium","Low")))`

✅ **Check:** 5 High, 9 Medium, 3 Low.

> **Try this:** *"What does this return for a merchant with TPV of exactly 50,000?"* The answer is **Medium**, because the test is `>=`.

---

## Exercise 5 — Explain a Formula

Copy the formula from **O3** and paste it into Copilot Chat:

```
Explain this formula in simple terms, part by part:
=IF(H3="","Not Installed",IF(N3>I3,"Breach","Within SLA"))
```

> **This works on any formula.** Try it on the most complicated formula in your own MIS.

---

## Exercise 6 — Pivot Table: Activation by Area

Copilot Chat **cannot build the pivot.** It gives you the click-path.

**Ask:**
```
Give me step-by-step instructions to build a PivotTable from A2:Q27 showing
the count of merchants for each Area, broken down by Activation Status
(column P). Column M is a blank spacer with no header — tell me how to handle it.
```

**Apply:**
1. Type `Spacer` into **M2** (a pivot needs a header in every column).
2. Select **A2:Q27** → **Insert → PivotTable → New Worksheet** → OK.
3. Drag the fields:
   - **Rows:** Area
   - **Columns:** Activation Status
   - **Values:** Merchant ID *(Count)*

✅ **Check:** Nashik has 2 Inactive. Mumbai Andheri has only 2 of 5 Active.

**No-pivot alternative:** `=COUNTIFS($D$3:$D$27,$S3,$P$3:$P$27,T$2)` puts areas down the side and statuses across the top.

---

## ⚠️ Data Handling — Read This

This is a **synthetic training dataset**, so it is safe to paste. Real merchant data is not:
- Use **Copilot Chat signed in with your BFL work account**.
- Never paste PAN, bank account, GST number, phone number, card details, OTP or UPI PIN.
- **De-identify first.** Replace names with M-1, M-2 and so on, and paste only the columns the question needs.

---

## What Changes With a Copilot Licence

| | **Copilot Chat** *(today)* | **Copilot in Excel** *(licensed)* |
|---|---|---|
| Sees your workbook | ✗ — you describe it | ✓ |
| Writes formulas | ✓ | ✓ |
| **Inserts** formulas | ✗ — you paste | ✓ |
| Builds PivotTables | ✗ — gives the steps | ✓ |
| Explains any formula | ✓ | ✓ |

---

*Bajaj Finance Limited | AI Academy — Payments (Sales), Field | Microsoft Copilot Training Programme*
*Training delivered by **Technizer India** | GenAI for Enterprise Teams*
