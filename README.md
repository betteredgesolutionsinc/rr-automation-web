# BESI Shift Assigner

This is a separate Apps Script web app bound to a **native Google Sheets** copy of the TK. The upload you linked is an Excel Office file in Drive, so first open it and use **File → Save as Google Sheets**. Keep the Excel original untouched.

1. From the new Google Sheets copy, open **Extensions → Apps Script**.
2. Paste `Code.gs` into the code file and create an HTML file named exactly `Index`, then paste `Index.html`.
3. Set the project time zone to **Asia/Manila**, or paste `appsscript.json` using Project Settings → Show manifest file.
4. Save and run `setupShiftSystem` once in the editor. Authorize access. This creates `SHIFT_OPTIONS` in the TK spreadsheet. If older versions already created `SHIFT_ASSIGNMENTS` and `SHIFT_ASSIGNMENT_HISTORY`, the new version no longer uses them.
5. Open `SHIFT_OPTIONS` and edit the sample shift names and hours to match your actual available shifts. Keep each Shift ID unique and set Status to `ACTIVE` for shifts to show in the app. Set it to `INACTIVE` to hide a shift. The sample options are 21:00–06:00, 00:00–09:00, 01:00–10:00, and 04:00–13:00.
6. Deploy → New deployment → Web app. Start with **Execute as: Me** and **Who has access: Only myself**. Open the web app URL.

As you type a name, matching employees appear with the available shifts. Click the employee and a shift, then **Process Today**. The app uses the server's current date in Asia/Manila and fills only that date's `TIME IN` and `TIME OUT` cells in the TK. It does not ask for a date range. These planned values are highlighted yellow and each cell gets a `BESI SCHEDULED SHIFT` note. It refuses to overwrite an existing value without that note. Today's date must already exist above a Time In/Time Out pair in one TK sheet. For example, if it is September 27 while the only TK sheet shows October 1–15, processing stops until a September 27 column or October's date arrives. Processing a midnight shift fills `00:00` and `09:00`; the early shift fills `01:00` and `10:00` when those are the configured shift hours.

The roster is read from sheets with an `EMPLOYEE NAME` or `COMPLETE NAME` heading. Nonblank employee IDs distinguish duplicate names; otherwise names are matched by normalized spelling. If you later need supervisors to assign shifts, add a role/access check before widening the web app deployment. A web app deployed to execute as the owner uses the owner's authority for every visitor.
