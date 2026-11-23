# CrossHope LightningStep tools

Browser helpers for CrossHope staff in LightningStep (chtc.lightningstep.com). No patient data is stored here.

## LS Forms (`ls-forms.user.js`)

Adds an **Open a form** button to the bottom right of every LightningStep page. Pick a form, type the patient's name or MRN, and the form opens on that patient's chart.

### Install (Chrome on Windows, once per computer)

1. Add **Tampermonkey** from the Chrome Web Store.
2. Right-click the Tampermonkey icon, choose **Manage extension**, and turn on **Allow User Scripts**.
3. On the same page, set **Site access** to **On specific sites** and add `https://chtc.lightningstep.com`.
4. Open [the install link](https://raw.githubusercontent.com/Mundles/crosshope-ls-tools/main/ls-forms.user.js) and click **Install**.
5. Open LightningStep. The **Open a form** button is at the bottom right.

### Updates

Tampermonkey checks this file for a higher `@version` and updates on its own. To get an update right away, click the Tampermonkey icon, then **Check for userscript updates**.

Only Mendy Teitlebaum approves changes to this repository.
