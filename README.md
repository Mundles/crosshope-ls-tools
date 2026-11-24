# CrossHope LightningStep tools

Browser helpers for CrossHope staff in LightningStep (chtc.lightningstep.com). No patient data is stored here.

## LS Forms (`ls-forms.user.js`)

Adds buttons to the bottom right of LightningStep pages:
- **Open a form**: pick a form, type the patient's name or MRN, and the form opens on that patient's chart.
- **New patient**: paste the Teams VOB, drop the prescreen PDF, check one screen, Save. Enters the inquiry, payer, admission and allergies, then opens the Biopsych.
- **New letter**: opens a fresh copy of the CrossHope letterhead (logo, header, footer) in Word for the web. Each click makes a new file in SharePoint > Restaura BH > Documents > Arkansas Clinic > CrossHope Letters. The master in `_Template (do not edit)` never changes. Needs access to that folder.
- **Treatment plans**: lists every PHP and IOP patient on the census who has no plan yet, whose review is overdue, or whose review is due this week, oldest first, with an Update plan now button on each row. The button turns red with a count when anything is missing or overdue. PHP reviews are counted every 20 weekdays, IOP every 30 days (settings at the top of the TP module).
- On a chart: **Ready for Assessment**, **Level of care** (step up/down, or fix a patient entered in the wrong program) and **Discharge**. Buttons: MH Biopsych, SUD Biopsych, LOCUS, ASAM, CIWA-Ar, COWS, C-SSRS + Safety Plan, Safety Plan, Individual Note, Treatment Plan Review.

### Install (Chrome on Windows, once per computer)

1. Add **Tampermonkey** from the Chrome Web Store.
2. Right-click the Tampermonkey icon, choose **Manage extension**, and turn on **Allow User Scripts**.
3. On the same page, set **Site access** to **On specific sites** and add `https://chtc.lightningstep.com`.
4. Open [the install link](https://raw.githubusercontent.com/Mundles/crosshope-ls-tools/main/ls-forms.user.js) and click **Install**.
5. Open LightningStep. The **Open a form** button is at the bottom right.

### Updates

Tampermonkey checks this file for a higher `@version` and updates on its own. To get an update right away, click the Tampermonkey icon, then **Check for userscript updates**.

Only Mendy Teitlebaum approves changes to this repository.
