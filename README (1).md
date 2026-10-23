# Sales order vs invoice report

A Streamlit dashboard that reads three Google Sheets – **Sale order**, **Sale Invoice** and
**Client_master** – and shows month-wise order, invoice and shortfall figures with a
Party Group ▸ Party ▸ Order-line drill-down. All heavy calculation happens in Python
(pandas), so nothing is computed in AppSheet virtual columns.

## How the numbers are calculated

| Item | Rule |
|---|---|
| Month | Taken from **SODATE** of the sales order (`20260701` → Jul-2026). Invoices inherit the month of their SO. |
| Link order ↔ invoice | `SODOCNO` + `ITEM NAME` (case, spaces and punctuation ignored). |
| Party group | Order `PARTY NAME` matched to Client master `Party Name` (GST no. used as fallback if both sheets have it). Unmatched parties show as **UNMAPPED PARTY** and are listed in *Data checks*. |
| Shortfall qty | `max(order qty − invoiced qty, 0)` per SO line. |
| Shortfall amount | Shortfall qty × order rate of that line. |
| Fill % | Invoiced ÷ ordered (qty or value). |
| Discontinued | Lines whose *Partial Status* contains "Discontinue" – shortfall on these is shown separately and can be excluded with a sidebar switch. |
| Gross weight | Read once per bill (it repeats on every line); packages are summed. |

## Reports

1. **Overview** – KPIs, month-wise order/invoice/shortfall chart and table, value by line status, top shortfall parties.
2. **Group ▸ Party drill-down** – click a party group → its parties → that party's SO lines.
3. **Month pivot** – any metric × month for party group, party, item, zone, state, salesman, platform or type.
4. **Items** – item-wise fill rate, ABC classification, top shortfall items, item × month heat-map.
5. **Order status & ageing** – fully / partly / not invoiced SOs, pending-order ageing, days from SO to first bill.
6. **Region & salesman** – zone, state, city, salesman, platform, direct/indirect, payment term.
7. **Dispatch** – bills, transporters, delivery locations, packages, weight, value per kg.
8. **Pricing & discount** – discount from MRP by group, invoice discounts, invoice rate vs order rate.
9. **Data checks** – unmapped parties, invoices with unknown SODOCNO, item-name mismatches, duplicates.

Every report follows the sidebar filters, and **Download Excel report** exports the filtered summaries.

## Deploy (GitHub + Streamlit Community Cloud, free)

1. Create a new GitHub repository (private is fine) and upload every file in this folder,
   including the hidden `.streamlit/config.toml`. Do **not** upload a real `secrets.toml`.
2. Go to <https://share.streamlit.io>, sign in with GitHub, click **Create app ▸ Deploy a public app from GitHub**.
3. Pick the repo, branch `main`, main file `app.py`, and choose a custom subdomain
   (e.g. `mycompany-sales-report`).
4. Open **Advanced settings ▸ Secrets**, paste the contents of `.streamlit/secrets.toml.example`
   with your values, and click **Deploy**.
5. Your report is live at `https://<subdomain>.streamlit.app`. Every `git push` redeploys it automatically.

### Giving the app access to the sheets

* **Option A – quick:** share each sheet as *Anyone with the link – Viewer* and leave out
  the `[gcp_service_account]` block. Anyone with the sheet link could then read it,
  including contact numbers in Client master.
* **Option B – private (recommended):** in Google Cloud Console create a project, enable
  the *Google Sheets API* and *Google Drive API*, create a **service account**, download its JSON key,
  paste the key into the `[gcp_service_account]` secrets block, and share the three sheets
  with the service account's `client_email` as *Viewer*.

Data is cached for 10 minutes; the **Refresh data** button reloads immediately.

## Run on your computer

```bash
pip install -r requirements.txt
streamlit run app.py
```
Without secrets the app opens on **Sample data** (in `sample_data/`), so you can explore every
report before connecting your sheets. You can also upload CSV/XLSX exports from the sidebar.
